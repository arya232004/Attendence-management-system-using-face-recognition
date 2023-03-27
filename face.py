import cv2 as cv
import tkinter as tk
import numpy as np
import pandas 
import time
import openpyxl
import face_recognition

root = tk.Tk()
c = tk.Canvas(root, bg="pink", height="200")
def check_fields(name_entry, roll_entry, submit_button, button_cam):
    # check if all input fields are filled
    if roll_entry.get() and name_entry.get():
        submit_button.config(state="normal")
        button_cam.config(state="normal")
    else:
        submit_button.config(state="disabled")
        button_cam.config(state="disabled")

def adduser():
    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)

    name = tk.Label(frame, text="Enter name of student: ")
    name.pack()

    name_entry = tk.Entry(frame)
    name_entry.pack()
    name_entry.bind("<KeyRelease>", lambda event: check_fields(name_entry, roll_entry, submit_button, button_cam))

    roll = tk.Label(frame, text="Enter Roll no: ")
    roll.pack()

    roll_entry = tk.Entry(frame)
    roll_entry.pack()
    roll_entry.bind("<KeyRelease>", lambda event: check_fields(name_entry, roll_entry, submit_button, button_cam))

    button_cam = tk.Button(frame, text="Take Photo", command=lambda: takephoto(name_entry.get(), roll_entry.get()), width=19, height=2, state="disabled")
    button_cam.pack()

    submit_button = tk.Button(frame, text="Submit", state="disabled")
    submit_button.pack()

    check_fields(name_entry, roll_entry, submit_button, button_cam)

    button = tk.Button(frame, text="Close", command=frame.destroy)
    button.pack()

def displaydata(frame,name,face_locations):
     for (top, right, bottom, left) in face_locations:
            cv.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
            cv.putText(frame, name, (left, top), cv.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
            

def attendence():
    df=pandas.read_excel("attendence.xlsx")
    known_encodings=[]
    name="xddd"
    roll=""
    dataframe = openpyxl.load_workbook("attendence.xlsx")
    dataframe1 = dataframe.active
    for row in range(1, dataframe1.max_row):
          li=[]
          stren=''
          for col in dataframe1.iter_cols(3, 3):
            stren=col[row].value
            for i in stren.split(","):
              li.append(float(i))  
            known_encodings.append(li)
    cam= cv.VideoCapture(0)    
    while True:
        ret, frame = cam.read()
        rgb_frame = frame[:, :, ::-1]
        face_locations = face_recognition.face_locations(rgb_frame,model="hog")
        if(len(face_locations)==1):
            live_encoding = face_recognition.face_encodings(frame, face_locations)
            if live_encoding:
                results = face_recognition.compare_faces(known_encodings, live_encoding[0])
                if True in results or name=="" or roll=="": 
                    print("Match Found")
                    name=dataframe1.cell(row+1, 2).value
                    roll=dataframe1.cell(row+1, 1).value 
                    displaydata(frame,name,face_locations)
                    # for (top, right, bottom, left) in face_locations:
                    #    cv.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
                    #    cv.putText(frame, name, (left, top), cv.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
                else:
                    pass
                    print("Match Not Found") 
                    name="unknown"
                    displaydata(frame,name,face_locations)
        elif(len(face_locations)>1):
            # cv.imshow('Video', np.zeros((480, 640, 3), np.uint8))
            print("No of faces should be 1 only")       
        else:
            # cv.imshow('Video', np.zeros((480, 640, 3), np.uint8))
            print("No face detected")
        cv.imshow('Video', frame)    
        
        if cv.waitKey(1) & 0xFF == ord('q'):
            cam.release()
            cv.destroyAllWindows()   
            break

    

def takephoto(name,roll):
    workb = openpyxl.load_workbook("attendence.xlsx")
    worksheet = workb.active
    last_row = worksheet.max_row
    face_encode_list=[]
    cam_port = 0
    cam = cv.VideoCapture(cam_port)
    time.sleep(2) 
    result, image = cam.read()
    print(result, image)
    if result:
        cv.imshow(name+roll+".jpg", image)
        cv.imwrite(name+roll+".jpg", image)
        
        imagefr = face_recognition.load_image_file(name+roll+".jpg")
        face_locations = face_recognition.face_locations(imagefr)
        face_encodings = face_recognition.face_encodings(imagefr, face_locations)
        print(face_encodings[0])
        for i in range(len(face_encodings[0])):
            print(face_encodings[0][1])
            face_encode_list.append(face_encodings[0][i])
        print(face_encode_list)    
        with open("Arya7.txt", "a") as file:
            for i in range(len(face_encodings[0])):
                file.write(str(face_encodings[0][i]))
                file.write("\n")
        if not  worksheet.cell(row=last_row, column=1).value:
            row_to_write = last_row
        else:
            row_to_insert = last_row + 1            
        worksheet.insert_rows(row_to_insert)    
        
        roll_no_cell = worksheet.cell(row=row_to_insert, column=1)
        name_cell = worksheet.cell(row=row_to_insert, column=2)
        face_encod=worksheet.cell(row=row_to_insert, column=3)

        name_cell.value = name
        roll_no_cell.value = roll 
        face_encod.value=','.join(map(str, face_encodings[0]))
        # face_encod.value = face_encodings[0]
        
        workb.save('attendence.xlsx')
        cv.waitKey(0)
        
        # df=df.append({"Name":name,"Roll":roll,},ignore_index=True)
        cv.destroyWindow("Attendence")
    else:
        print("No image detected. Please! try again")
    c.pack()
    root.mainloop()


root.title("Attendence")
root.geometry("500x500")
root.configure(bg="white")
btn = tk.Button(root, text="Add User", fg="red", command=adduser, width=19, height=2)
btn.place(x=180, y=40)

btn = tk.Button(root, text="Attendence", fg="red", command=attendence, width=19, height=2)
btn.place(x=180, y=150)
root.mainloop()
