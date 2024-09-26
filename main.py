import sys
import unohelper
import officehelper
from com.sun.star.task import XJobExecutor

try:
    import speech_recognition as sr
except ModuleNotFoundError as e:
    print(f"Error importing speech_recognition: {e}")

#import speech_recognition as sr
from enum import Enum

class Language(Enum):
    ENGLISH = "en-US"
    CHINESE = "zh-TW"
    JAPANESE = "ja-JP"

class SpeechToText():
    def print_mic_device_index():
        for index, name in enumerate(sr.Microphone
        .list_microphone_names()):
            print("{1}, device_index={0}".format
            (index, name))

    def speech_to_text(device_index,
    language=Language.ENGLISH):
       print("Entering speech_to_text method")
       # Initialize the recognizer 
       r = sr.Recognizer()
       with sr.Microphone(device_index=device_index) as source:
            print("Start Talking:")
            audio = r.listen(source, timeout=5)  # 设置超时为 5 秒

            # 檢查錄製的音訊是否為空
            if audio.frame_data:
                print("Audio recorded successfully.")
            else:
                print("No audio recorded.")
                return "No audio recorded."

            try:
                text = r.recognize_google(audio,
                language=language.value)
                print("You said: {}".format(text))
                return text  # 返回識別的文本

            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
                return "Recognition failed."
            except sr.RequestError as e:
                print(f"Could not request results from Google Speech Recognition service; {e}")
                return "Recognition failed."
            except Exception as e:
                print(f"Error recognizing speech: {e}")
                return "Recognition failed."

            #except:
                #print("Please try again.")

class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        # handling different situations (inside LibreOffice or other process)
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)
            
    def trigger(self, args):
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", self.ctx)
        model = desktop.getCurrentComponent()
        if not hasattr(model, "Text"):
            model = self.desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        text = model.Text
        cursor = text.createTextCursor()
        device_index = 1  # 根據你的麥克風設置更改這個值
        recognized_text = SpeechToText.speech_to_text(device_index=device_index)
        text.insertString(cursor, "Start :" + recognized_text + "\n", 0)
        #text.insertString(cursor, "Hello Extension argument -> " + args + "\n", 0)

# Starting from Python IDE
def main():
    try:
        ctx = XSCRIPTCONTEXT
        print("XSCRIPTCONTEXT is available")
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
        else:
            print("Office successfully bootstrapped")
    job = MainJob(ctx)
    job.trigger("hello")


# Starting from command line
if __name__ == "__main__":
    main()

# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    MainJob,  # UNO object class
    "org.extension.stt.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)

# vim: set shiftwidth=4 softtabstop=4 expandtab:        