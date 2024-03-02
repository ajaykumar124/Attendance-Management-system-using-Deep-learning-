from tkinter import *
from PIL import Image, ImageTk
import cv2
import face_recognition
import os

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance System")

        # Open the image using PIL and convert it to a Tkinter-compatible format
        pil_image = Image.open("bg3.jpg")
        self.background_image = ImageTk.PhotoImage(pil_image)
        
        background_label = Label(root, image=self.background_image)
        background_label.place(relwidth=1, relheight=1)  # Cover the entire window

        self.cap = cv2.VideoCapture(0)
        self.images = []
        self.names = []

        # GUI Elements
        self.label = Label(root)
        self.name_var = StringVar()
        self.name_entry = Entry(root, textvariable=self.name_var, font=('calibre', 12, 'normal'), justify='center')
        self.name_entry.insert(0, 'Enter Name')  # Placeholder text
        self.update_btn = Button(root, text='Update', command=self.updatedata, font=('calibre', 12, 'bold'), bg='#4CAF50', fg='white')
        self.snap_btn = Button(root, text='Snapshot', command=self.snapshot, font=('calibre', 12, 'bold'), bg='#008CBA', fg='white')
        self.quit_btn = Button(root, text='Quit', command=self.quitapp, font=('calibre', 12, 'bold'), bg='#FF0000', fg='white')

        # Call the function to display frames
        self.show_frames()

        # Place elements in the center
        self.label.place(relx=0.5, rely=0.3, anchor=CENTER)
        self.name_entry.place(relx=0.5, rely=0.5, anchor=CENTER)
        self.update_btn.place(relx=0.3, rely=0.7, anchor=CENTER, width=150, height=50)
        self.snap_btn.place(relx=0.5, rely=0.7, anchor=CENTER, width=150, height=50)
        self.quit_btn.place(relx=0.7, rely=0.7, anchor=CENTER, width=150, height=50)

    def snapshot(self):
        ret, frame = self.cap.read()
        if not ret:
            print("Failed to grab frame")
            return
        cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(cv2image)
        face_locations = face_recognition.face_locations(cv2image)
    
        for face_location in face_locations:
            top, right, bottom, left = face_location
            im1 = img.crop((left, top, right, bottom))
            name = self.name_var.get()
            if not name:
                print("Name is empty")
                return
            image_path = os.path.join(r'D:\attendancesystem\images', f'{name}.jpg')
            im1.save(image_path)

    def updatedata(self):
        images_directory = r"D:\attendancesystem\images"
        for file_name in os.listdir(images_directory):
            file_path = os.path.join(images_directory, file_name)
            if os.path.isfile(file_path):
                image = cv2.imread(file_path)
                b = os.path.splitext(file_name)[0]
                self.names.append(b)
                self.images.append(image)
                print(self.names)

    def show_frames(self):
        ret, frame = self.cap.read()
        if not ret:
            print("Failed to grab frame")
            return
        cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(cv2image)
        imgtk = ImageTk.PhotoImage(image=img)
        self.label.imgtk = imgtk
        self.label.configure(image=imgtk)
        self.label.after(10, self.show_frames)

    def quitapp(self):
        self.root.destroy()
        self.cap.release()

if __name__ == "__main__":
    root = Tk()
    root.attributes('-fullscreen', True)  # Open in fullscreen
    app = AttendanceApp(root)
    root.mainloop()
