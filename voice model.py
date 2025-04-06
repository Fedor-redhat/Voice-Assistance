import os
import json
import queue
import threading
import webbrowser
import customtkinter as ctk
import speech_recognition as sr
import pyttsx3
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.naive_bayes import MultinomialNB
import pyautogui
import win32gui
import win32process
import psutil

class VoiceAssistant:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Голосовой Ассистент Pro")
        self.root.geometry("800x600")
        
        self.command_queue = queue.Queue()
        self.custom_commands_file = "commands.json"
        self.custom_commands = []
        self.is_listening = False
        self.always_listen = False
        
        self.setup_gui()
        self.load_custom_commands()
        self.setup_ai_model()
        self.setup_voice_engine()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.check_microphone)
        self.root.after(100, self.process_queue)
        self.root.mainloop()

    def on_close(self):
        self.always_listen = False
        self.root.destroy()

    def check_microphone(self):
        try:
            if not sr.Microphone().list_microphone_names():
                self.show_error("Микрофон не обнаружен!")
        except Exception as e:
            self.log(f"Ошибка микрофона: {str(e)}")

    def setup_voice_engine(self):
        try:
            self.engine = pyttsx3.init()
            self.engine.setProperty('rate', 150)
            self.recognizer = sr.Recognizer()
            self.microphone = sr.Microphone()
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
        except Exception as e:
            self.log(f"Ошибка инициализации: {str(e)}")

    def setup_ai_model(self):
        try:
            base_texts = [
                "выключи компьютер", "открой проводник", "найди в интернете",
                "открой меню пуск", "перезагрузи систему", "открой настройки",
                "хватит", "стоп", "перестань слушать", 
                "слушай", "начинай работать", "активируйся",
                "закрой браузер", "закрой проводник", "закрой приложение"
            ]
            base_intents = [
                "shutdown", "open_explorer", "web_search",
                "open_start_menu", "reboot", "open_settings",
                "deactivate", "deactivate", "deactivate",
                "activate", "activate", "activate",
                "close_app", "close_app", "close_app"
            ]
            
            self.vectorizer = CountVectorizer()
            texts = base_texts + [cmd['phrase'] for cmd in self.custom_commands]
            intents = base_intents + [cmd['intent'] for cmd in self.custom_commands]
            
            X = self.vectorizer.fit_transform(texts)
            self.model = MultinomialNB()
            self.model.fit(X, intents)
        except Exception as e:
            self.log(f"Ошибка модели: {str(e)}")

    def process_queue(self):
        while not self.command_queue.empty():
            try:
                task = self.command_queue.get_nowait()
                task()
            except Exception as e:
                self.log(f"Ошибка очереди: {str(e)}")
        self.root.after(100, self.process_queue)

    def setup_gui(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        self.log_text = ctk.CTkTextbox(self.main_frame, wrap="word")
        self.log_text.pack(pady=10, padx=10, fill="both", expand=True)
        
        control_frame = ctk.CTkFrame(self.main_frame)
        control_frame.pack(pady=10)
        
        self.listen_button = ctk.CTkButton(
            control_frame,
            text="🎤 Начать слушать",
            command=lambda: self.command_queue.put(self.toggle_listening)
        )
        self.listen_button.pack(side="left", padx=5)
        
        ctk.CTkButton(
            control_frame,
            text="➕ Добавить команду",
            command=lambda: self.command_queue.put(self.show_add_command_dialog)
        ).pack(side="left", padx=5)

    def show_add_command_dialog(self):
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Новая команда")
        dialog.geometry("400x200")
        
        entries = {}
        fields = [
            ("Фраза для активации", "phrase"),
            ("Действие (путь или запрос)", "action")
        ]
        
        for text, key in fields:
            ctk.CTkLabel(dialog, text=text).pack(pady=5)
            entries[key] = ctk.CTkEntry(dialog, width=300)
            entries[key].pack()
        
        def save_command():
            phrase = entries['phrase'].get().strip()
            action = entries['action'].get().strip()
            if phrase and action:
                self.custom_commands.append({
                    'phrase': phrase,
                    'action': action,
                    'intent': f"custom_{len(self.custom_commands)+1}"
                })
                self.save_custom_commands()
                self.setup_ai_model()
                dialog.destroy()
        
        ctk.CTkButton(dialog, text="Сохранить", command=save_command).pack(pady=10)

    def load_custom_commands(self):
        try:
            if os.path.exists(self.custom_commands_file):
                with open(self.custom_commands_file, 'r') as f:
                    self.custom_commands = json.load(f)
            else:
                self.custom_commands = []
        except Exception as e:
            self.log(f"Ошибка загрузки команд: {str(e)}")
            self.custom_commands = []

    def save_custom_commands(self):
        try:
            with open(self.custom_commands_file, 'w') as f:
                json.dump(self.custom_commands, f, indent=2)
        except Exception as e:
            self.log(f"Ошибка сохранения команд: {str(e)}")

    def toggle_listening(self):
        if not self.is_listening:
            self.start_listening()
        else:
            self.stop_listening()

    def start_listening(self):
        self.is_listening = True
        self.always_listen = True
        self.listen_button.configure(text="🔴 Слушаю...")
        self.speak("Готов к работе")
        threading.Thread(target=self.safe_listen_loop, daemon=True).start()

    def stop_listening(self):
        self.is_listening = False
        self.always_listen = False
        self.listen_button.configure(text="🎤 Начать слушать")
        self.speak("Режим ожидания")

    def safe_listen_loop(self):
        try:
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
            
            while self.always_listen and self.is_listening:
                try:
                    with self.microphone as source:
                        audio = self.recognizer.listen(
                            source, 
                            timeout=2,
                            phrase_time_limit=4
                        )
                    text = self.recognizer.recognize_google(
                        audio, 
                        language="ru-RU",
                        show_all=False
                    )
                    self.command_queue.put(lambda: self.process_command(text))
                except sr.WaitTimeoutError:
                    continue
                except sr.UnknownValueError:
                    self.command_queue.put(lambda: self.log("Не разобрал речь"))
                except Exception as e:
                    self.command_queue.put(lambda: self.log(f"Ошибка: {str(e)}"))
        except Exception as e:
            self.command_queue.put(lambda: self.log(f"Ошибка цикла: {str(e)}"))

    def process_command(self, text):
        try:
            self.log(f"Вы: {text}")
            intent = self.predict_intent(text.lower())
            
            if intent == "activate":
                if not self.is_listening:
                    self.start_listening()
                return
            elif intent == "deactivate":
                if self.is_listening:
                    self.stop_listening()
                return
                
            self.execute_command(intent, text)
        except Exception as e:
            self.log(f"Ошибка обработки: {str(e)}")

    def predict_intent(self, text):
        try:
            X_new = self.vectorizer.transform([text])
            return self.model.predict(X_new)[0]
        except Exception as e:
            return "unknown"

    def execute_command(self, intent, text):
        try:
            system_commands = {
                "shutdown": lambda: os.system("shutdown /s /t 1"),
                "reboot": lambda: os.system("shutdown /r /t 1"),
                "open_explorer": self.open_explorer,
                "open_start_menu": lambda: pyautogui.press('win'),
                "open_settings": lambda: os.system("start ms-settings:"),
                "web_search": lambda: webbrowser.open(f"https://google.com/search?q={text[16:]}"),
                "close_app": lambda: self.close_application(text)
            }
            
            for cmd in self.custom_commands:
                if cmd['intent'] == intent:
                    self.handle_custom_command(cmd['action'])
                    return
            
            if intent in system_commands:
                system_commands[intent]()
                self.log(f"Выполнено: {intent}")
            else:
                self.log("Команда не распознана")
        except Exception as e:
            self.log(f"Ошибка выполнения: {str(e)}")

    def open_explorer(self):
        try:
            os.startfile(os.path.expanduser("~"))
        except Exception as e:
            self.log(f"Ошибка проводника: {str(e)}")

    def close_application(self, text):
        try:
            app_name = text.lower().replace("закрой", "").strip()
            apps = {
                "браузер": "chrome.exe",
                "проводник": "explorer.exe",
                "блокнот": "notepad.exe",
                "приложение": self.get_foreground_process()
            }
            
            process_name = apps.get(app_name)
            if process_name:
                os.system(f"taskkill /f /im {process_name}")
                self.log(f"Закрываю {app_name}")
            else:
                self.log(f"Приложение не найдено: {app_name}")
        except Exception as e:
            self.log(f"Ошибка закрытия: {str(e)}")

    def get_foreground_process(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return psutil.Process(pid).name()
        except:
            return "unknown.exe"

    def handle_custom_command(self, action):
        try:
            if action.endswith(".exe"):
                os.startfile(action)
            elif action.startswith(("http://", "https://")):
                webbrowser.open(action)
            else:
                os.system(action)
            self.log(f"Выполнено: {action}")
        except Exception as e:
            self.log(f"Ошибка команды: {str(e)}")

    def speak(self, text):
        def safe_speak():
            try:
                if not self.engine._inLoop:
                    self.engine = pyttsx3.init()
                self.engine.say(text)
                self.engine.runAndWait()
            except Exception as e:
                self.log(f"Ошибка синтеза: {str(e)}")
        threading.Thread(target=safe_speak, daemon=True).start()

    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def show_error(self, message):
        self.log_text.insert("end", f"ОШИБКА: {message}\n")

if __name__ == "__main__":
    assistant = VoiceAssistant()