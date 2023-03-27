import cv2 as cv
import face_recognition as fr
cam_port = 0
cam = cv.VideoCapture(cam_port)


result, image = cam.read()


if result:

	cv.imshow("Attendence", image)
	cv.imwrite("GeeksForGeeks.png", image)
	cv.waitKey(0)
	cv.destroyWindow("GeeksForGeeks")
else:
	print("No image detected. Please! try again")
