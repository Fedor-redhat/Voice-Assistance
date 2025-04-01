import os
import webbrowser
import speech_recognition as sr
import pyttsx3
import pyautogui
import subprocess

class VoiceAssistant:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.engine = pyttsx3.init()
        self.commands = {
            'открой браузер': self.open_browser,
            'выключи компьютер': self.shutdown_pc,
            'перезагрузи компьютер': self.reboot_pc,
            'закрой программу': self.close_program,
            'скрой':self.minimize_window,
            'разверни':self.full_window,
            'переключи окно':self.rolling,
            'список': self.list_commands
        }
        
        # Настройки голосового движка для русского языка
        voices = self.engine.getProperty('voices')
        for voice in voices:
            if 'russian' in voice.languages:
                self.engine.setProperty('voice', voice.id)
                break

    def listen(self):
        with sr.Microphone() as source:
            print("Слушаю...")
            self.recognizer.adjust_for_ambient_noise(source)
            audio = self.recognizer.listen(source)
            
            try:
                text = self.recognizer.recognize_google(audio, language="ru-RU").lower()
                print("Распознано:", text)
                return text
            except Exception as e:
                print("Ошибка распознавания:", e)
                return ""

    def speak(self, text):
        print("Ассистент>>", text)
        self.engine.say(text)
        self.engine.runAndWait()
    
    def minimize_window(self):
        pyautogui.hotkey('win', 'down')
        self.speak("окно скрыто")

    def full_window(self):
        pyautogui.hotkey('win', 'up')
        self.speak("окно открыто")

    def open_browser(self):
        webbrowser.open('https://www.google.com')
        self.speak("Браузер открыт")

    def rolling(self):
        pyautogui.hotkey('win', 'control', 'right')
        self.speak("окно переключено")

    def shutdown_pc(self):
        self.speak("Выключаю компьютер через 1 минуту")
        os.system("shutdown /s /t 60")

    def reboot_pc(self):
        self.speak("Перезагружаю компьютер через 1 минуту")
        os.system("shutdown /r /t 60")

    def close_program(self):
        pyautogui.hotkey('alt', 'f4')
        self.speak("Программа закрыта")

    def list_commands(self):
        commands_list = ", ".join(self.commands.keys())
        self.speak(f"Доступные команды: {commands_list}")

    def handle_command(self, command):
        for key in self.commands:
            if key in command:
                self.commands[key]()
                return True
        return False

    def run(self):
        self.speak("Приветствую, ассистент запущен. Чтобы узнать список команд скажите 'список'")
        while True:
            command = self.listen()
            
            if not command:
                continue
                
            if 'стоп' in command:
                self.speak("Выключаюсь")
                break
                
            if not self.handle_command(command):
                self.speak("Не распознала команду. Попробуйте еще раз")

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.run()