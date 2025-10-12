# --*utf-8*--
# Author: 一念断星河（修订：单线程复用引擎 + 线程内 COM + 打包兜底）
import threading, queue, ctypes
import pyttsx3

# --- 确保 PyInstaller 收到 SAPI 驱动 ---
try:
    import pyttsx3.drivers  # noqa
    import pyttsx3.drivers.sapi5  # noqa
except Exception:
    pass

# --- COM 初始化：pythoncom -> comtypes -> ctypes/ole32 ---
def _co_init():
    # COINIT_APARTMENTTHREADED = 0x2
    COINIT_APARTMENTTHREADED = 0x2
    try:
        import pythoncom
        pythoncom.CoInitialize()
        return ("pythoncom", None)
    except Exception:
        try:
            import comtypes
            try:
                comtypes.CoInitialize()
            except Exception:
                # 有些版本暴露的是 CoInitializeEx
                from comtypes import CoInitializeEx  # type: ignore
                CoInitializeEx(0, COINIT_APARTMENTTHREADED)
            return ("comtypes", None)
        except Exception:
            try:
                ole32 = ctypes.OleDLL("ole32")
                ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
                return ("ctypes", ole32)
            except Exception as e:
                print("TTS COM init failed:", e)
                return (None, None)

def _co_uninit(tag_ole32):
    tag, ole32 = tag_ole32
    try:
        if tag == "pythoncom":
            import pythoncom
            pythoncom.CoUninitialize()
        elif tag == "comtypes":
            import comtypes
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
        elif tag == "ctypes" and ole32 is not None:
            try:
                ole32.CoUninitialize()
            except Exception:
                pass
    except Exception:
        pass


class Voice:
    def __init__(self):
        # 可被 UI 修改的状态
        self._lock = threading.RLock()
        self.m_Switch = True
        self.m_volume = 1.0   # 0.0~1.0 或 0~100（自动识别）
        self.m_rate   = 230   # 接受 wpm 或 0~100 百分比

        # 运行态
        self._q = queue.Queue()
        self._stop = threading.Event()
        self._worker = None

        # 存档（可选）
        try:
            from core import core_save
            from core.core_define import Path_Voice
            self._core_save = core_save
            self._path_voice = Path_Voice
        except Exception:
            self._core_save = None
            self._path_voice = None

    # -------- 外部接口 --------
    def InitVoice(self, volume=None, rate=None):
        # 读取历史配置（与原版兼容）
        try:
            if self._core_save and self._path_voice:
                cfg = self._core_save.LoadJson(self._path_voice) or {}
                if volume is None:
                    volume = cfg.get("volume", None)
                if rate is None:
                    rate = cfg.get("rate", None)
                sw = cfg.get("switch", None)
                if sw is not None:
                    self.Switch(bool(sw))
        except Exception:
            pass

        with self._lock:
            if volume is not None:
                self._set_volume_nolock(volume)
            if rate is not None:
                self._set_rate_nolock(int(rate))
        self._ensure_worker()

    def Speak(self, text: str):
        if not text or not self.m_Switch:
            return
        self._ensure_worker()
        self._q.put(str(text))

    def Switch(self, bSwitch: bool):
        with self._lock:
            self.m_Switch = bool(bSwitch)
        self.Save()

    def SetVolume(self, vol):
        with self._lock:
            self._set_volume_nolock(vol)
        self.Save()

    def SetRate(self, rate_wpm_or_pct: int):
        with self._lock:
            self._set_rate_nolock(int(rate_wpm_or_pct))
        self.Save()

    def SetRatePercent(self, pct: int):
        with self._lock:
            self.m_rate = max(0, min(100, int(pct)))
        self.Save()

    def Save(self):
        if not (self._core_save and self._path_voice):
            return
        try:
            self._core_save.SaveJson(self._path_voice, {
                "rate":   int(self.m_rate),
                "volume": float(self.m_volume),
                "switch": bool(self.m_Switch),
            })
        except Exception:
            pass

    def Stop(self):
        try:
            while not self._q.empty():
                self._q.get_nowait()
        except Exception:
            pass

    # -------- 内部：归一化/映射 --------
    def _set_rate_nolock(self, v: int):
        self.m_rate = v  # 可是 0~100 或 wpm

    def _set_volume_nolock(self, v):
        try:
            v = float(v)
            if v > 1.0:  # 支持 0~100
                v = v / 100.0
        except Exception:
            v = 1.0
        self.m_volume = max(0.0, min(1.0, v))

    def _map_rate_to_wpm(self, v: int) -> int:
        # [-10..10] 相对；[0..100] 百分比；否则视作 wpm
        if -10 <= v <= 10:
            return 220 + v * 12
        if 0 <= v <= 100:
            return int(80 + (v / 100.0) * 300)  # 80~380
        return max(80, min(450, int(v)))

    # -------- 工作线程 --------
    def _ensure_worker(self):
        if self._worker and self._worker.is_alive():
            return
        self._stop.clear()
        t = threading.Thread(target=self._run, name="TTSWorker", daemon=True)
        t.start()
        self._worker = t

    def _run(self):
        tag = _co_init()
        engine = None

        def _ensure_engine():
            nonlocal engine
            if engine is None:
                # 先 sapi5，失败退回默认
                try:
                    engine = pyttsx3.init(driverName="sapi5")
                except Exception:
                    engine = pyttsx3.init()

        try:
            while not self._stop.is_set():
                try:
                    text = self._q.get(timeout=0.2)
                except queue.Empty:
                    continue
                if not text:
                    continue

                try:
                    _ensure_engine()
                    # 每次播报前应用“当前”设置
                    with self._lock:
                        vol = float(self.m_volume)
                        rate_wpm = self._map_rate_to_wpm(int(self.m_rate))

                    try: engine.setProperty("volume", vol)
                    except Exception: pass
                    try: engine.setProperty("rate", rate_wpm)
                    except Exception: pass

                    engine.say(text)
                    engine.runAndWait()

                except Exception as e:
                    print("TTS run error:", e)
                    try:
                        if engine is not 无:
                            engine.stop()
                    except Exception:
                        pass
                    engine = 无  # 下次重建

        finally:
            try:
                if engine is not 无:
                    engine.stop()
            except Exception:
                pass
            _co_uninit(tag)

    def __del__(self):
        try:
            self._stop.set()
            self._q.put(None)
            if self._worker and self._worker.is_alive():
                self._worker。join(timeout=1.0)
        except Exception:
            pass


# 单例 & 兼容旧接口
if "g_Instance" not in globals():
    g_Instance = Voice()

Initialize       = g_Instance.InitVoice
Speak            = g_Instance.Speak
SetRate          = g_Instance.SetRate
SetRatePercent   = g_Instance.SetRatePercent
SetVolume        = g_Instance.SetVolume
