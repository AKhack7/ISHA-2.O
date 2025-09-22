def toggle_voice(self):
    if self.microphone is None:
        self.is_listening = False
        self.voice_button.configure(style='Voice.Off.TButton')
        message = "Microphone unavailable, turning off."
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")
        time.sleep(1)
        self.is_listening = True
        self.voice_button.configure(style='Voice.On.TButton')
        message = "Trying to turn microphone back on."
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")
        return
    self.is_listening = not self.is_listening
    if self.is_listening:
        self.voice_button.configure(style='Voice.On.TButton')
        message = "Microphone is now on"
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")
        threading.Thread(target=self.listen_voice, daemon=True).start()
    else:
        self.voice_button.configure(style='Voice.Off.TButton')
        message = "Microphone is now off"
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")








def listen(self):
    if self.microphone is None:
        self.is_listening = False
        self.voice_button.configure(style='Voice.Off.TButton')
        message = "Microphone unavailable, turning off for 1 second."
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")
        time.sleep(1)
        self.is_listening = True
        self.voice_button.configure(style='Voice.On.TButton')
        message = "Trying to turn microphone back on."
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")
        return None
    try:
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            for _ in range(3):  # 3 रिट्रीज़
                try:
                    audio = self.recognizer.listen(source, timeout=10, phrase_time_limit=10)
                    return self.recognizer.recognize_google(audio).lower()
                except sr.WaitTimeoutError:
                    message = "No speech detected, retrying..."
                    self.speak(message)
                    self.chat_box_insert(f"Output: {message}\n")
                    continue
                except sr.UnknownValueError:
                    message = "Could not understand audio, retrying..."
                    self.speak(message)
                    self.chat_box_insert(f"Output: {message}\n")
                    continue
                except sr.RequestError as e:
                    message = f"Speech recognition service error: {str(e)}. Retrying..."
                    self.speak(message)
                    self.chat_box_insert(f"Output: {message}\n")
                    continue
            # सभी रिट्री फेल होने पर
            self.is_listening = False
            self.voice_button.configure(style='Voice.Off.TButton')
            message = "Voice input failed after retries, turning off for 1 second."
            self.speak(message)
            self.chat_box_insert(f"Output: {message}\n")
            time.sleep(1)
            self.is_listening = True
            self.voice_button.configure(style='Voice.On.TButton')
            message = "Trying to turn microphone back on."
            self.speak(message)
            self.chat_box_insert(f"Output: {message}\n")
            return None
    except Exception as e:
        self.is_listening = False
        self.voice_button.configure(style='Voice.Off.TButton')
        message = f"Voice input failed: {str(e)}. Turning off for 1 second."
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")
        time.sleep(1)
        self.is_listening = True
        self.voice_button.configure(style='Voice.On.TButton')
        message = "Trying to turn microphone back on."
        self.speak(message)
        self.chat_box_insert(f"Output: {message}\n")
        return None
