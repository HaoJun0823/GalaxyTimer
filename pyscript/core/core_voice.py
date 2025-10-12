# Desc: 稳定的 TTS：单线程队列 + 线程内 COM 初始化，避免跨线程复用 pyttsx3 引擎

import threading
import queue
import pyttsx3

from core.core_define import Path_Voice  # 如需读设置可保留
try:
    import pythoncom as _pycom
    def _co_init():
        _pycom.CoInitialize()
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
        # 与旧接口兼容的字段
        self.m_Switch = True
        self.m_volume = 1.0
        self.m_rate = 200

        # 内部状态
        self._q = queue.Queue()
        self._stop = threading.Event()
        self._worker = None

    # 供外部调用：初始化（可选）
    def InitVoice(self, volume: float = None, rate: int = None):
        if volume is not None: self.m_volume = float(volume)
        if rate   is not None: self.m_rate   = int(rate)
        self._ensure_worker()

    # 供外部调用：说话
    def Speak(self, text: str):
        if not self.m_Switch or not text:
            return
        self._ensure_worker()
        # 将文本与当前音量/语速一起入队，确保每条播报自带配置
        self._q.put((str(text), float(self.m_volume), int(self.m_rate)))

    # 供外部调用：停止/清队（不会销毁线程）
    def Stop(self):
        try:
            while not self._q.empty():
                self._q.get_nowait()
        except Exception:
            pass

    # -------- 内部：工作线程 --------
    def _ensure_worker(self):
        if self._worker and self._worker.is_alive():
            return
        self._stop.clear()
        t = threading.Thread(target=self._run, name="TTSWorker", daemon=True)
        t.start()
        self._worker = t

    def _run(self):
        _co_init()  # 关键：在线程内初始化 COM
        try:
            while not self._stop.is_set():
                try:
                    item = self._q.get(timeout=0.1)
                except queue.Empty:
                    continue
                if item is None:
                    continue
                text, vol, rate = item

                engine = None
                try:
                    # 明确使用 SAPI5；失败则退回默认
                    try:
                        engine = pyttsx3.init(driverName="sapi5")
                    except Exception:
                        engine = pyttsx3.init()

                    try: engine.setProperty("volume", float(vol))
                    except Exception: pass
                    try: engine.setProperty("rate",   int(rate))
                    except Exception: pass

                    engine.say(text)
                    engine.runAndWait()
                except Exception as e:
                    print("语音播报异常:", e)
                finally:
                    if engine is not None:
                        try: engine.stop()
                        except Exception: pass
                        # 立刻丢弃 engine，避免跨事件复用引擎引发的状态机问题
                        del engine
        finally:
            _co_uninit()

    # 退出时清理
    def __del__(self):
        try:
            self._stop.set()
            self._q.put(None)
            if self._worker and self._worker.is_alive():
                self._worker.join(timeout=1.0)
        except Exception:
            pass


# 单例导出（与原接口保持一致）
if "g_Instance" not in globals():
    g_Instance = Voice()

Initialize = g_Instance.InitVoice
Speak      = g_Instance.Speak
