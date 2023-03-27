import cv2
import tkinter as tk
from PIL import Image, ImageTk

# Use OpenCV to capture video from the default camera
video_capture = cv2.VideoCapture(0)

# Create a tkinter window
root = tk.Tk()

# Create a canvas to display the video feed
canvas = tk.Canvas(root, width=640, height=480, bg="white")
canvas.pack()

def show_video():
    # Read a frame from the video feed
    ret, frame = video_capture.read()

    if not ret:
        print("Error: Failed to capture video frame")
        return

    # Convert the image from BGR color (which OpenCV uses) to RGB color (which tkinter uses)
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    # Create a PIL Image from the OpenCV image
    try:
        image = Image.fromarray(rgb_frame)
    except Exception as e:
        print("Error: Failed to create PIL Image:", e)
        return

    # Display the image on the canvas
    try:
        photo = ImageTk.PhotoImage(image)
        canvas.create_image(0, 0, image=photo, anchor=tk.NW)
    except Exception as e:
        print("Error: Failed to create PhotoImage:", e)
        return

    # Repeat the function every 10 milliseconds
    root.after(10, show_video)

# Call the function to start showing the video
show_video()

# Start the tkinter main loop
root.mainloop()

# Release the camera
video_capture.release()
