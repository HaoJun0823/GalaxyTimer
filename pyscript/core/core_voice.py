# --*utf-8*--
# 安全的语音播报：后台线程 + 队列，Speak() 不阻塞 & 不抛异常
import threading, queue

try:
    import pythoncom
    import win32com.client as win32
except Exception:  # 打包缺组件也不让外面崩
    pythoncom = 无
    win32 = 无

# 全局队列与线程
_speak_q: "queue.Queue[str|None]" = queue.Queue()
_worker: threading.Thread | 无 = 无
_default_rate: int = 0  # 语速：SAPI -10..10，0 为默认
_default_volume: int = 100  # 音量 0..100

def _worker_loop():
    # 在工作线程里初始化 COM，避免阻塞/干扰主线程
    if pythoncom:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
    voice = 无
    try:
        if win32:
            try:
                voice = win32.Dispatch("SAPI.SpVoice")
                try:
                    voice.Rate = _default_rate
                except Exception:
                    pass
                try:
                    voice.Volume = _default_volume
                except Exception:
                    pass
            except Exception:
                voice = 无

        while True:
            text = _speak_q.get()
            if text is 无:
                break
            try:
                if voice:
                    # 同步播报，但在后台线程里，不阻塞调用方
                    voice.Speak(str(text))
            except Exception:
                # 吃掉播报异常，避免影响热键/UI
                pass
    finally:
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

def _ensure_worker():
    global _worker
    if _worker is None or not _worker.is_alive():
        _worker = threading.Thread(target=_worker_loop, name="VoiceWorker", daemon=True)
        _worker.start()

def Speak(text: str):
    """非阻塞：把文本丢到后台线程播放；任何异常都不会向外冒出。"""
    try:
        _ensure_worker()
        _speak_q.put_nowait(str(text))
    except Exception:
        pass  # 绝不影响调用方

def SetRate(rate: int):
    """设置语速（-10..10），对后续播报生效。"""
    global _default_rate
    try:
        _default_rate = int(rate)
    except Exception:
        _default_rate = 0

def SetVolume(vol: int):
    """设置音量（0..100），对后续播报生效。"""
    global _default_volume
    try:
        _default_volume = max(0, min(100, int(vol)))
    except Exception:
        _default_volume = 100

def Shutdown():
    """可选：程序退出时调用，优雅结束语音线程。"""
    try:
        _speak_q.put_nowait(None)
    except Exception:
        pass
