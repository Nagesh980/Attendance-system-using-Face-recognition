import cv2
import pickle
import numpy as np
import os
from tkinter import *
from tkinter import messagebox
import threading
import yagmail
from sklearn.neighbors import KNeighborsClassifier
import csv
import time
from datetime import datetime, date
from win32com.client import Dispatch
import schedule

# Global variables
LABELS = []
FACES = []

# Function to speak text
def speak(text):
    speak_engine = Dispatch("SAPI.SpVoice")
    speak_engine.Speak(text)

# Function to load existing data
def load_data():
    global LABELS, FACES
    if os.path.exists('data/names.pkl') and os.path.exists('data/faces_data.pkl'):
        with open('data/names.pkl', 'rb') as w:
            LABELS = pickle.load(w)
        with open('data/faces_data.pkl', 'rb') as f:
            FACES = pickle.load(f)

# Function to save data
def save_data():
    with open('data/names.pkl', 'wb') as w:
        pickle.dump(LABELS, w)
    with open('data/faces_data.pkl', 'wb') as f:
        pickle.dump(FACES, f)

# Function to collect face data
def collect_data(name):
    if not name:
        messagebox.showerror("Error", "Please enter your name.")
        return
    
    video = cv2.VideoCapture(0)
    facedetect = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
    faces_data = []
    i = 0

    while True:
        try:
            ret, frame = video.read()
            if not ret:
                print("Failed to capture image from camera.")
                break
            
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = facedetect.detectMultiScale(gray, 1.3, 5)
            
            for (x, y, w, h) in faces:
                crop_img = frame[y:y+h, x:x+w, :]
                resized_img = cv2.resize(crop_img, (50, 50))
                if len(faces_data) < 1000 and i % 10 == 0:
                    faces_data.append(resized_img)
                i += 1
                
                # Draw green frame and progress bar
                cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
                bar_length = 300  # Width of progress bar
                progress = len(faces_data) / 100
                progress_bar_width = 400
                progress_bar_height = 30
                bar_x = int((frame.shape[1] - progress_bar_width) / 2)
                bar_y = frame.shape[0] - 50
                cv2.rectangle(frame, (bar_x, bar_y), 
                              (bar_x + progress_bar_width, bar_y + progress_bar_height), 
                              (255, 255, 255), 2)
                cv2.rectangle(frame, (bar_x, bar_y), 
                              (bar_x + int(progress_bar_width * progress), bar_y + progress_bar_height), 
                              (0, 255, 0), -1)
            
            cv2.imshow("Frame", frame)
            k = cv2.waitKey(1)
            if len(faces_data) == 100:
                break
        except Exception as e:
            print(f"An error occurred: {e}")
            break

    video.release()
    cv2.destroyAllWindows()

    faces_data = np.asarray(faces_data)
    faces_data = faces_data.reshape(100, -1)

    # Ensure data directory exists
    if 'data' not in os.listdir():
        os.mkdir('data')

    global LABELS, FACES
    if not os.path.exists('data/names.pkl'):
        LABELS = [name] * 100
        FACES = faces_data
    else:
        with open('data/names.pkl', 'rb') as w:
            LABELS = pickle.load(w)
        with open('data/faces_data.pkl', 'rb') as f:
            FACES = pickle.load(f)
        
        LABELS += [name] * 100
        FACES = np.append(FACES, faces_data, axis=0)

    save_data()
    messagebox.showinfo("Info", "Data collected successfully!")

# Function to run attendance system
def run_attendance_system():
    video = cv2.VideoCapture(0)
    facedetect = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

    with open('data/names.pkl', 'rb') as w:
        LABELS = pickle.load(w)
    with open('data/faces_data.pkl', 'rb') as f:
        FACES = pickle.load(f)

    knn = KNeighborsClassifier(n_neighbors=5)
    knn.fit(FACES, LABELS)

    COL_NAMES = ['NAME', 'TIME']
    logged_names_today = set()  # Set to track who has been marked present today
    current_date = date.today()  # Get the current date
    unknown_threshold = 4000  # Distance threshold for unknown classification

    # Define color codes
    GREEN = (0, 255, 0)  # Bright Green
    RED = (0, 0, 255)    # Bright Red

    while True:
        ret, frame = video.read()
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = facedetect.detectMultiScale(gray, 1.3, 5)

        # Reset logged_names_today if the date has changed
        if date.today() != current_date:
            logged_names_today.clear()
            current_date = date.today()

        for (x, y, w, h) in faces:
            crop_img = frame[y:y+h, x:x+w, :]
            resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
            resized_img = resized_img.reshape(-1, 7500)
            distances, indices = knn.kneighbors(resized_img)
            distance = distances[0][0]
            output = LABELS[indices[0][0]] if distance < unknown_threshold else "Unknown"

            ts = time.time()
            timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
            exist = os.path.isfile(f"Attendance/Attendance_{current_date}.csv")

            # Set color based on whether the face is recognized or unknown
            frame_color = GREEN if output != "Unknown" else RED
            text_color = (0, 0, 0)  # Black text for better contrast

            # Draw the face rectangle
            cv2.rectangle(frame, (x, y), (x+w, y+h), frame_color, 2)
            cv2.rectangle(frame, (x, y-40), (x+w, y), frame_color, -1)
            cv2.putText(frame, output, (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 0.75, text_color, 1)

            # Mark attendance and speak
            if output != "Unknown" and output not in logged_names_today:
                logged_names_today.add(output)
                attendance = [output, timestamp]

                if exist:
                    with open(f"Attendance/Attendance_{current_date}.csv", "a", newline='') as csvfile:
                        writer = csv.writer(csvfile)
                        writer.writerow(attendance)
                else:
                    with open(f"Attendance/Attendance_{current_date}.csv", "w", newline='') as csvfile:
                        writer = csv.writer(csvfile)
                        writer.writerow(COL_NAMES)
                        writer.writerow(attendance)

                speak(f"Attendance Taken for {output}")

            elif output == "Unknown":
                speak("Unknown person detected")

        cv2.imshow("Frame", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    video.release()
    cv2.destroyAllWindows()


# Function to send email with attendance file
def send_email():
    email_address = "nageshkumbar70@gmail.com"
    password = "knnw jzxm draw xgwz"
    receiver_email = "Nageshsk453@gmail.com"
    
    yag = yagmail.SMTP(email_address, password)
    current_date = datetime.now().strftime("%Y-%m-%d")
    file_path = f"Attendance/Attendance_{current_date}.csv"
    
    if os.path.exists(file_path):
        subject = f"Attendance for {current_date}"
        body = "Please find the attached attendance file."
        yag.send(to=receiver_email, subject=subject, contents=body, attachments=file_path)
        print("Email sent successfully.")
    else:
        print("No attendance file found for today.")

# Function to schedule email sending
def schedule_email():
    schedule.every().day.at("10:16").do(send_email)
    while True:
        schedule.run_pending()
        time.sleep(60)

# Function to view registered users
def view_users():
    global LABELS
    if LABELS:
        users = "\n".join(set(LABELS))
        messagebox.showinfo("Registered Users", users)
    else:
        messagebox.showinfo("No Users", "No users registered yet.")

# GUI Setup
def start_gui():
    def collect_data_thread():
        name = name_entry.get()
        if not name:
            messagebox.showerror("Error", "Please enter your name.")
        else:
            threading.Thread(target=collect_data, args=(name,), daemon=True).start()

    # GUI Setup
    root = Tk()
    root.title("Face Recognition Attendance System")
    root.geometry("500x400")
    root.configure(bg="#f0f0f0")

    label = Label(root, text="Face Recognition Attendance System", font=("Arial", 20), bg="#f0f0f0")
    label.pack(pady=20)

    # Name entry
    name_label = Label(root, text="Enter Name:", font=("Arial", 14), bg="#f0f0f0")
    name_label.pack(pady=5)
    name_entry = Entry(root, font=("Arial", 14))
    name_entry.pack(pady=5)

    # Collect Data
    collect_button = Button(root, text="Collect Data", command=collect_data_thread, font=("Arial", 14), bg="#2196F3", fg="#ffffff")
    collect_button.pack(pady=10)

    # View Registered Users
    view_button = Button(root, text="View Registered Users", command=view_users, font=("Arial", 14), bg="#2196F3", fg="#ffffff")
    view_button.pack(pady=10)

    # Start Attendance System
    start_button = Button(root, text="Start Attendance System", command=lambda: threading.Thread(target=run_attendance_system, daemon=True).start(), font=("Arial", 14), bg="#2196F3", fg="#ffffff")
    start_button.pack(pady=10)

    # Start Email Scheduling
    threading.Thread(target=schedule_email, daemon=True).start()

    root.mainloop()

# Run GUI
if __name__ == "__main__":
    load_data()
    start_gui()
