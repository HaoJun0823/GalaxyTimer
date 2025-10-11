# pyscript/core/core_voice.py  —— 在文件顶部补充这几行
import threading
import pyttsx3
from core import core_save
from core.core_define import Path_Voice

# 兼容两种 COM 初始化方式（优先 pythoncom，没有就用 comtypes）
try:
    import pythoncom
    def _co_init(): pythoncom.CoInitialize()
    def _co_uninit(): pythoncom.CoUninitialize()
except Exception:
    try:
        import comtypes
        def _co_init(): comtypes.CoInitialize()
        def _co_uninit(): comtypes.CoUninitialize()
    except Exception:
        def _co_init(): pass
        def _co_uninit(): pass

class Voice:
    def __init__(self):
        # 如果原来有 self._engine，建议删掉跨线程复用，改为“每次新建引擎”（线程更安全）
        self._cur_speak_text = None

    def InitVoice(self):
        # 如有需要可以加载保存的音量/语速配置
        pass

    def _speak_worker(self, text: str):
        _co_init()
        try:
            # 每次在本线程创建引擎，避免跨线程共享
            engine = pyttsx3.init()  # Windows 默认就是 sapi5
            # 从配置加载参数（若没有就走默认）
            data = core_save.LoadJson(Path_Voice) if Path_Voice else {}
            vol = data.get("volume", 1.0)
            rate = data.get("rate", 200)
            engine.setProperty("volume", vol)
            engine.setProperty("rate", rate)

            engine.say(text)
            engine.runAndWait()
            engine.stop()
        finally:
            _co_uninit()

    def Speak(self, text: str):
        # 丢到后台线程，避免阻塞 UI，也规避调用方所在线程的 COM 状态
        t = threading.Thread(target=self._speak_worker, args=(text,), daemon=True)
        t.start()

# 单例保持不变
if "g_Instance" not in globals():
    g_Instance = Voice()

Initialize = g_Instance.InitVoice
Speak = g_Instance.Speak

import threading

import pyttsx3

from core import core_save, core_timer
from core.core_define import Path_Voice


class Voice:
    def __init__(self):
        self._engine = None
        self._cur_speak_text = None
        self.m_Switch = True
        self.m_volume = 1
        self.m_rate = 200

        self._Timer_Volume = None

    def InitVoice(self):
        try:
            engine = pyttsx3.init()
        except:
            try:
                engine = pyttsx3.init("sapi5")
            except:
                try:
                    engine = pyttsx3.init("dummy")
                except:
                    return
        if not engine:
            print("当前系统不支持语音")
            self._engine = None
            self.Switch(False)
            self.Save()
            return
        self._engine = engine
        voice_config = core_save.LoadJson(Path_Voice)
        self.SetRate(voice_config.get("rate", 200))
        self.SetVolume(voice_config.get("volume", 1))
        self.Switch(voice_config.get("switch", True))
        # 音色
        voices = self._engine.getProperty('voices')
        print("当前系统支持的音色：")
        for _voice in voices:
            print(_voice.name, "  id:", _voice.id)
        if not voices:
            print("当前系统不支持语音")
            self._engine = None
            self.Switch(False)
            self.Save()
            return
        self._engine.setProperty('voice', voices[0].id)

    def Switch(self, bSwitch):
        self.m_Switch = bSwitch

    def SetVolume(self, volume):
        if self._cur_speak_text:
            core_timer.CreateOnceTimer(20, self.SetVolume, delta=False, volume=volume)
            return
        print("设置音量:", volume)
        self.m_volume = volume
        if not self._engine:
            return
        self._engine.setProperty('volume', self.m_volume)  # 音量 0-1

    def SetRate(self, rate):
        if self._cur_speak_text:
            core_timer.CreateOnceTimer(20, self.SetRate, delta=False, rate=rate)
            return
        print("设置语速:", rate)
        self.m_rate = rate
        if not self._engine:
            return
        self._engine。setProperty('rate', self.m_rate)  # 语速 0-200

    def Speak(self, text):
        if not self.m_Switch:
            return
        if not self._engine:
            return
        if self._cur_speak_text:
            core_timer.CreateOnceTimer(20, self.Speak, delta=False, text=text)
            return
        self._cur_speak_text = text
        threading.Thread(target=self.Real_Speak, daemon=True).start()

    def Real_Speak(self):
        if self._engine and self._cur_speak_text:
            self._engine.say(self._cur_speak_text)
            self._engine.runAndWait()
        self._cur_speak_text = None

    def Save(self):
        save_data = {
            "rate": self.m_rate,
            "volume": self.m_volume，
            "switch": self.m_Switch
        }
        core_save.SaveJson(Path_Voice, save_data)

    def Stop(self):
        if self._engine:
            try:
                self._engine.stop()
            except:
                pass

    def __del__(self):
        self.Stop()


if "g_Instance" not 在 globals():
    g_Instance = Voice()

Initialize = g_Instance.InitVoice
Speak = g_Instance.Speak
