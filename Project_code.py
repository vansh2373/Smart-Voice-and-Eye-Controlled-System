import os
import keyboard
import cv2
import dlib
import numpy as np
import pyautogui
import speech_recognition as sr
import threading
from pynput.mouse import Controller
import tkinter as tk
import queue
import subprocess
import win32api
import win32con
import win32gui

# Global vars
screen_width, screen_height = pyautogui.size()
pyautogui.FAILSAFE = False
detector = dlib.get_frontal_face_detector()
predictor = dlib.shape_predictor("shape_predictor_68_face_landmarks.dat")
mouse = Controller()
recognizer = sr.Recognizer()
mic_index = 0  # Change if needed
status_queue = queue.Queue()

# Folders
SPECIAL_FOLDERS = {
    "this pc": "explorer.exe shell:MyComputerFolder",
    "recycle bin": "explorer.exe shell:RecycleBinFolder",
    "documents": "explorer shell:Personal",
    "downloads": "explorer shell:Downloads",
    "desktop": "explorer shell:Desktop"
}

LEFT_EYE_POINTS = [36, 37, 38, 39, 40, 41]
RIGHT_EYE_POINTS = [42, 43, 44, 45, 46, 47]

def get_eye_landmarks(landmarks, eye_points):
    return np.array([(landmarks.part(point).x, landmarks.part(point).y) for point in eye_points], dtype=np.int32)

def get_eye_center(eye):
    return np.mean(eye, axis=0).astype("int")

def recognize_command():
    while True:
        try:
            with sr.Microphone(device_index=mic_index) as source:
                status_queue.put("🎙️ Listening...")
                recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = recognizer.listen(source, timeout=3)

            command = recognizer.recognize_google(audio).lower()
            print(f"Recognized: {command}")
            status_queue.put(f"✅ {command}")

            if "one click" in command:
                pyautogui.click()
            elif "double click" in command:
                pyautogui.doubleClick()
            elif "right click" in command:
                pyautogui.rightClick()
            elif "scroll up" in command:
                pyautogui.scroll(500)
            elif "scroll down" in command:
                pyautogui.scroll(-500)
            elif "open" in command:
                folder_name = command.replace("open", "").strip()
                if folder_name in SPECIAL_FOLDERS:
                    os.system(SPECIAL_FOLDERS[folder_name])
                else:
                    os.system(f'explorer "{folder_name}"')
            elif "close" in command:
                pyautogui.hotkey("alt", "f4")
            elif "minimize" in command:
                pyautogui.hotkey("win", "down")
            elif "maximize" in command:
                pyautogui.hotkey("win", "up")
            elif "volume up" in command:
                pyautogui.press("volumeup")
            elif "volume down" in command:
                pyautogui.press("volumedown")
            elif "mute" in command:
                pyautogui.press("volumemute")
            elif "switch tab" in command:
                if "next" in command:
                    pyautogui.hotkey("ctrl", "tab")
                elif "previous" in command:
                    pyautogui.hotkey("ctrl", "shift", "tab")

        except sr.WaitTimeoutError:
            status_queue.put("⏳ Waiting...")
        except sr.UnknownValueError:
            status_queue.put("🤔 Didn't catch that")
        except Exception as e:
            print("Speech error:", e)
            status_queue.put("⚠️ Mic Error")

def run_eye_tracker():
    cap = cv2.VideoCapture(0)
    prev_x, prev_y = 0, 0
    smoothing = 0.2

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        frame_height, frame_width = frame.shape[:2]
        box_size = int(min(frame_width, frame_height) * 0.1)
        box_x1, box_y1 = (frame_width - box_size) // 2, (frame_height - box_size) // 2
        box_x2, box_y2 = box_x1 + box_size, box_y1 + box_size

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = detector(gray)

        for face in faces:
            landmarks = predictor(gray, face)
            left_eye = get_eye_landmarks(landmarks, LEFT_EYE_POINTS)
            right_eye = get_eye_landmarks(landmarks, RIGHT_EYE_POINTS)
            left_center = get_eye_center(left_eye)
            right_center = get_eye_center(right_eye)
            eye_center = (left_center + right_center) // 2

            eye_x = np.clip(eye_center[0], box_x1, box_x2)
            eye_y = np.clip(eye_center[1], box_y1, box_y2)

            target_x = np.interp(eye_x, [box_x1, box_x2], [screen_width, 0])
            target_y = np.interp(eye_y, [box_y1, box_y2], [0, screen_height])

            prev_x = prev_x * (1 - smoothing) + target_x * smoothing
            prev_y = prev_y * (1 - smoothing) + target_y * smoothing
            mouse.position = (prev_x, prev_y)

            cv2.rectangle(frame, (box_x1, box_y1), (box_x2, box_y2), (255, 0, 0), 2)
            cv2.circle(frame, tuple(left_center), 3, (0, 255, 0), -1)
            cv2.circle(frame, tuple(right_center), 3, (0, 255, 0), -1)
            cv2.circle(frame, tuple(eye_center), 3, (255, 0, 0), -1)

        cv2.imshow("Eye Tracker", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    status_queue.put("EXIT_APP")

def run_status_overlay():
    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.attributes("-transparentcolor", "black")
    root.configure(bg="black")
    root.geometry(f"250x50+{screen_width - 270}+30")

    label = tk.Label(root, text="🎙️ Ready", font=("Helvetica", 14),
                     fg="white", bg="black")
    label.pack(expand=True, fill="both")

    def update():
        while not status_queue.empty():
            msg = status_queue.get()
            if msg == "EXIT_APP":
                root.quit()
                return
            label.config(text=msg)
        root.after(100, update)

    update()
    root.mainloop()

# === START ===
if __name__ == "__main__":
    threading.Thread(target=recognize_command, daemon=True).start()
    threading.Thread(target=run_eye_tracker, daemon=True).start()
    run_status_overlay()
