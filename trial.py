import face_recognition

# load the input image
image = face_recognition.load_image_file("Arya7.jpg")

# find all the faces in the image
face_locations = face_recognition.face_locations(image)
face_encodings = face_recognition.face_encodings(image, face_locations)

# print the first face encoding
print(face_encodings[0])
print(face_encodings[0][2])


    
    
