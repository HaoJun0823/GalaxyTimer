# --*utf-8*--
# Author：一念断星河
# Crete Data：2025/10/12
# Desc：谁能想到这音频引擎修的这么坐牢...

import threading
import pyttsx3

from core import core_save, core_timer
from core.core_define import Path_Voice

# ---------------- COM 初始化（线程内） ----------------
try:
    import pythoncom  # 优先 pythoncom（pywin32）

    def _co_init():
        pythoncom.CoInitialize()

    def _co_uninit():
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

except Exception:
    try:
        import comtypes  # 其次 comtypes

        def _co_init():
            comtypes.CoInitialize()

        def _co_uninit():
            try:
                comtypes.CoUninitialize()
            except Exception:
                pass
    except Exception:
        def _co_init():
            pass

        def _co_uninit():
            pass


class Voice:
    def __init__(self):
        # 和原逻辑保持同名字段，避免 UI/菜单侧耦合出问题
        self._engine = None             # 不常驻复用，仅保留字段以兼容旧代码
        self._cur_speak_text = None
        self.m_Switch = True
        self.m_volume = 1.0
        self.m_rate = 200

        self._Timer_Volume = None  # 兼容：说话中延后设置

    # ---------------- 配置 ----------------
    def InitVoice(self):
        # 不在主线程创建 engine，避免跨线程使用
        voice_config = core_save.LoadJson(Path_Voice)
        self.SetRate(voice_config.get("rate", 200))
        self.SetVolume(voice_config.get("volume", 1.0))
        self.Switch(voice_config.get("switch", True))
        # 这里不探测 voices、也不绑定具体 voice，避免在主线程触碰 COM

    def Save(self):
        data = {
            "rate": int(self.m_rate),
            "volume": float(self.m_volume),
            "switch": bool(self.m_Switch),
        }
        core_save.SaveJson(Path_Voice, data)

    # ---------------- 参数 ----------------
    def Switch(self, bSwitch: bool):
        self.m_Switch = bool(bSwitch)

    def SetVolume(self, volume: float):
        if self._cur_speak_text:
            # 兼容原逻辑：说话中延后设置
            core_timer.CreateOnceTimer(20, self.SetVolume, delta=False, volume=volume)
            return
        try:
            self.m_volume = max(0.0, min(1.0, float(volume)))
        except Exception:
            pass

    def SetRate(self, rate: int):
        if self._cur_speak_text:
            core_timer.CreateOnceTimer(20, self.SetRate, delta=False, rate=rate)
            return
        try:
            r = int(rate)
        except Exception:
            return
        # pyttsx3 一般 0~500；此处不强钳，只作保护
        self.m_rate = max(-1000, min(1000, r))

    # ---------------- 播报 ----------------
    def Speak(self, text: str):
        if not self.m_Switch:
            return
        if not text:
            return
        if self._cur_speak_text:
            # 兼容原逻辑：队列化，避免并发
            core_timer.CreateOnceTimer(20, self.Speak, delta=False, text=text)
            return
        self._cur_speak_text = str(text)
        threading.Thread(target=self.Real_Speak, daemon=True).start()

    def _create_engine(self):
        """在当前线程创建 engine：优先 sapi5，失败降级 dummy；失败返回 None"""
        try:
            return pyttsx3.init("sapi5")
        except Exception:
            try:
                return pyttsx3.init("dummy")
            except Exception:
                return None

    def Real_Speak(self):
        _co_init()
        try:
            engine = self._create_engine()
            if not engine:
                print("当前系统不支持语音")
                return

            # 设置参数（容错）
            try:
                engine.setProperty("volume", float(self.m_volume))
            except Exception:
                pass
            try:
                engine.setProperty("rate"， int(self.m_rate))
            except Exception:
                pass

            engine.say(self._cur_speak_text)
            engine.runAndWait()
        except Exception as e:
            # 尽量打印，便于排查
            print("语音播报异常：", e)
        finally:
            try:
                # 个别后端需要显式 stop
                engine.stop()
            except Exception:
                pass
            self._cur_speak_text = 无
            _co_uninit()

    def Stop(self):
        # 兼容旧接口：当前实现为即说即弃，无需集中 stop
        pass

    def __del__(self):
        try:
            self.Stop()
        except Exception:
            pass


# 单例导出（与原接口保持一致）
if "g_Instance" not in globals():
    g_Instance = Voice()

Initialize = g_Instance.InitVoice
Speak = g_Instance.Speak
