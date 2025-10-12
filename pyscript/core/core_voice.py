# --*utf-8*--
import threading, queue
import pyttsx3

# 若项目里有设置路径依赖，可保留；没有也不影响
try:
    from core.core_define import Path_Voice  # noqa
except Exception:
    Path_Voice = None

# COM 初始化：优先 pythoncom，其次 comtypes，最后空实现
try:
    import pythoncom as _pycom
    def _co_init(): _pycom.CoInitialize()
    def _co_uninit():
        try: _pycom.CoUninitialize()
        except Exception: pass
except Exception:
    try:
        import comtypes
        def _co_init():
            try: comtypes.CoInitialize()
            except Exception: pass
        def _co_uninit():
            try: comtypes.CoUninitialize()
            except Exception: pass
    except Exception:
        def _co_init(): pass
        def _co_uninit(): pass


class Voice:
    def __init__(self):
        # —— 可供外部 / UI 修改的状态（加锁保护）——
        self._lock = threading.RLock()
        self.m_Switch = True
        self.m_volume = 1.0   # 0.0~1.0 或 0~100（自动识别）
        self.m_rate   = 230   # 推荐默认略快点；可传 wpm 或 0~100 百分比

        # —— 内部运行态 ——
        self._q = queue.Queue()
        self._stop = threading.Event()
        self._worker = None

    # ----- 外部接口 -----
    def InitVoice(self, volume: float | int | None = None, rate: int | None = None):
        with self._lock:
            if volume is not None: self._set_volume_nolock(volume)
            if rate   is not None: self._set_rate_nolock(rate)
        self._ensure_worker()

    def Speak(self, text: str):
        if not text or not self.m_Switch:
            return
        self._ensure_worker()
        # 只入队文本；音量/语速在消费时读取“最新值”
        self._q.put(str(text))

    # —— 新增给 UI 用的设置方法（线程安全）——
    def SetRate(self, rate_wpm_or_pct: int):
        with self._lock:
            self._set_rate_nolock(rate_wpm_or_pct)

    def SetRatePercent(self, pct: int):
        """UI 是 0~100 滑块时可用"""
        with self._lock:
            self.m_rate = max(0, min(100, int(pct)))

    def SetVolume(self, vol: float | int):
        with self._lock:
            self._set_volume_nolock(vol)

    def Stop(self):
        # 清空队列，不杀线程；播报会自然停止
        try:
            while not self._q.empty():
                self._q.get_nowait()
        except Exception:
            pass

    # ----- 内部：设置归一化 -----
    def _set_rate_nolock(self, v: int):
        # 支持三种输入：[-10..10]（相对）、[0..100]（百分比）、[80..450]（wpm）
        v = int(v)
        self.m_rate = v

    def _set_volume_nolock(self, v: float | int):
        try:
            v = float(v)
            # 若是 0~100，转为 0.0~1.0
            if v > 1.0: v = v / 100.0
        except Exception:
            v = 1.0
        self.m_volume = max(0.0, min(1.0, v))

    def _map_rate_to_wpm(self, v: int) -> int:
        # 输入 v 可能是：[-10..10]（相对）、[0..100]（百分比）、或已是 wpm
        if -10 <= v <= 10:
            # 相对档位：中心 220 wpm，每档 12 wpm
            return 220 + v * 12
        if 0 <= v <= 100:
            # 百分比：80~380 wpm 线性映射
            return int(80 + (v / 100.0) * 300)
        # 其他认为已经是 wpm
        return max(80, min(450, int(v)))

    # ----- 工作线程 -----
    def _ensure_worker(self):
        if self._worker and self._worker.is_alive():
            return
        self._stop.clear()
        t = threading.Thread(target=self._run, name="TTSWorker", daemon=True)
        t.start()
        self._worker = t

    def _run(self):
        _co_init()
        engine = None

        def _ensure_engine():
            nonlocal engine
            if engine is None:
                try:
                    engine = pyttsx3.init(driverName="sapi5")
                except Exception:
                    engine = pyttsx3.init()

        try:
            while not self._stop.is_set():
                try:
                    text = self._q.get(timeout=0.1)
                except queue.Empty:
                    continue
                if not text:
                    continue

                try:
                    _ensure_engine()
                    # 每次播报前，取“当前”设置并应用
                    with self._lock:
                        vol = float(self.m_volume)
                        rate_wpm = self._map_rate_to_wpm(int(self.m_rate))

                    try: engine.setProperty("volume", vol)
                    except Exception: pass
                    try: engine.setProperty("rate",   rate_wpm)
                    except Exception: pass

                    engine.say(text)
                    engine.runAndWait()

                except Exception as e:
                    print("TTS run error:", e)
                    # 出错就丢弃引擎，自动重建
                    try:
                        if engine is not None:
                            engine.stop()
                    except Exception:
                        pass
                    engine = None

        finally:
            try:
                if engine is not None:
                    engine.stop()
            except Exception:
                pass
            _co_uninit()

    def __del__(self):
        try:
            self._stop.set()
            self._q.put(None)
            if self._worker and self._worker.is_alive():
                self._worker.join(timeout=1.0)
        except Exception:
            pass


# 单例 & 兼容旧接口
if "g_Instance" not in globals():
    g_Instance = Voice()

Initialize = g_Instance.InitVoice
Speak      = g_Instance.Speak

# —— 可选：如果你的 UI 已经在其他地方引用了这些名字，也兼容提供 —— 
SetRate        = g_Instance.SetRate
SetRatePercent = g_Instance.SetRatePercent
SetVolume      = g_Instance.SetVolume
