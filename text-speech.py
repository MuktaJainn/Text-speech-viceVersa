import tkinter as tk
import win32com.client
import speech_recognition as sr


def speak():
    text = entry.get()
    speaker.Speak(text)


def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)

    try:
        recognized_text = recognizer.recognize_google(audio)
        entry.delete(0, tk.END)  # Clear the entry
        entry.insert(tk.END, recognized_text)
    except sr.UnknownValueError:
        print("Speech Recognition could not understand the audio")
    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")


# main window
root = tk.Tk()
root.title("Communication Aid App")
root.geometry("900x500+300+200")
root.configure(bg="#f0f0f0")  # Set background color

# frame for better organization
frame = tk.Frame(root, bg="#f0f0f0")
frame.pack(pady=20)

# label
label = tk.Label(frame, text="Enter text to speak:", font=("Arial", 16), bg="#f0f0f0")
label.grid(row=0, column=0, pady=10, padx=10)


entry_font = ("Arial", 20)
entry = tk.Entry(frame, width=50, font=entry_font)
entry.grid(row=1, column=0, pady=10, padx=10)

# buttons for speech-to-text and text-to-speech
listen_button = tk.Button(frame, text="1. Listen", command=listen, bg="#2196F3", fg="white", height=2, width=10,
                          font=("Arial", 14))
listen_button.grid(row=2, column=0, pady=10, padx=10)

speak_button = tk.Button(frame, text="2. Speak", command=speak, bg="#4caf50", fg="white", height=2, width=10,
                         font=("Arial", 14))
speak_button.grid(row=3, column=0, pady=10, padx=10)

# Create a button to close the application
quit_button = tk.Button(frame, text="Quit", command=root.destroy, bg="#f44336", fg="white", height=2, width=10,
                        font=("Arial", 14))
quit_button.grid(row=4, column=0, pady=10, padx=10)

speaker = win32com.client.Dispatch("SAPI.SpVoice")


root.mainloop()
