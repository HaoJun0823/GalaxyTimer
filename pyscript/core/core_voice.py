# --*utf-8*--
# Author：一念断星河
# Crete Data：2024/6/15
# Desc：不过是大梦一场空，不过是孤影照惊鸿。

import threading
from typing import Optional, Set

import pyttsx3

from core import core_save, core_timer  # 保持原有依赖
from core.core_define import Path_Voice


# 线程内的 COM 初始化（优先 pythoncom，没有就用 comtypes，实在没有就空实现）
try:
    import pythoncom

    def _co_init():
        pythoncom.CoInitialize()

    def _co_uninit():
        pythoncom.CoUninitialize()
except Exception:
    try:
        import comtypes

        def _co_init():
            comtypes.CoInitialize()

        def _co_uninit():
            comtypes.CoUninitialize()
    except Exception:
        def _co_init():
            pass

        def _co_uninit():
            pass


class Voice:
    def __init__(self):
        # 不再跨线程持有同一个 engine；改为“每次 Speak 新建引擎”
        self._enable: bool = True
        self._volume: float = 1.0
        self._rate: int = 400  # 与仓库中 voice.json 的默认值保持一致
        self._cur_speak_text: Optional[str] = None

        # 跟踪活动引擎，便于 Stop() 统一打断
        self._engines: Set[pyttsx3.Engine] = set()
        self._lock = threading.Lock()

    # ---------------- 配置读写 ----------------

    def Load(self):
        try:
            data = core_save.LoadJson(Path_Voice)
            self._enable = bool(data.get("switch", True))
            self._volume = float(data.get("volume", 1.0))
            self._rate = int(data.get("rate", 400))
        except Exception as e:
            print("读取语音配置失败：", e)

    def Save(self):
        try:
            data = {
                "switch": bool(self._enable),
                "volume": float(self._volume),
                "rate": int(self._rate),
            }
            core_save.SaveJson(Path_Voice, data)
        except Exception as e:
            print("保存语音配置失败：", e)

    def InitVoice(self):
        # 兼容旧接口：初始化时仅加载配置；不提前创建 engine（避免跨线程复用问题）
        self.Load()

    # ---------------- 参数设置 ----------------

    def Switch(self, bEnable: bool):
        self._enable = bool(bEnable)

    def SetVolume(self, vol: float):
        # 期望范围 0.0 ~ 1.0
        try:
            vol = float(vol)
        except Exception:
            return
        self._volume = max(0.0, min(1.0, vol))

    def SetRate(self, rate: int):
        # pyttsx3 的典型区间大致 0~500；这里不强制钳制，只做合理范围保护
        try:
            rate = int(rate)
        except Exception:
            return
        self._rate = max(-1000, min(1000, rate))

    # ---------------- 播报核心 ----------------

    def _register_engine(self, engine: pyttsx3.Engine):
        with self._lock:
            self._engines.add(engine)

    def _unregister_engine(self, engine: pyttsx3.Engine):
        with self._lock:
            if engine in self._engines:
                self._engines.remove(engine)

    def _speak_worker(self, text: str, volume: float, rate: int):
        _co_init()
        engine = None
        try:
            # Windows 优先 SAPI5，失败再降级 dummy
            try:
                engine = pyttsx3.init("sapi5")
            except Exception:
                try:
                    engine = pyttsx3.init("dummy")
                except Exception:
                    engine = 无

            if not engine:
                print("当前系统不支持语音")
                return

            self._register_engine(engine)

            # 设置参数
            try:
                engine.setProperty("volume", volume)
            except Exception:
                pass
            try:
                engine.setProperty("rate", rate)
            except Exception:
                pass

            engine.say(text)
            engine.runAndWait()
        except Exception as e:
            # 大多数 pyttsx3/SAPI 的问题会静默，这里尽量打印，便于排查
            print("语音播报异常：", e)
        finally:
            try:
                if engine:
                    engine.stop()
            except Exception:
                pass
            if engine:
                self._unregister_engine(engine)
            _co_uninit()

    def Speak(self, text: str):
        if not self._enable:
            return
        if not text:
            return
        # 读取当前配置快照，避免线程里再次读文件
        v, r = self._volume, self._rate
        t = threading.Thread(
            target=self._speak_worker, args=(str(text), v, r), daemon=True
        )
        t.start()

    def Stop(self):
        # 停止所有正在播报的 engine
        with self._lock:
            engines = list(self._engines)
            self._engines.clear()
        for e 在 engines:
            try:
                e.stop()
            except Exception:
                pass

    def __del__(self):
        self.Stop()


# 单例导出（兼容原有用法：core_voice.Speak / core_voice.Initialize）
if "g_Instance" not in globals():
    g_Instance = Voice()

Initialize = g_Instance.InitVoice
Speak = g_Instance.Speak
