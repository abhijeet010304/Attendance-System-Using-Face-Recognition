import numpy
import cv2
import pickle
import openpyxl
from datetime import date

attendance_sheet = openpyxl.load_workbook('Attendance.xlsx')
today = date.today()
month_string = today.strftime("%B")
sheet = attendance_sheet[month_string]

roll_generator = sheet[1]
roll_numbers = {}
day_generator = sheet['A']
day_numbers = {}

col_num = 2
row_num = 2

for x in roll_generator:
	if x.value != None:
		roll_numbers[x.value] = col_num
		col_num += 1 
for x in day_generator:
	if x.value != None:
		if len(str(x.value).split()) == 2:
			day_numbers[x.value.split()[0]] = row_num
			row_num += 1

# print(roll_numbers)
# print(day_numbers)

face_xml = "face.xml"
face_cascade = cv2.CascadeClassifier(face_xml)

recognizer = cv2.face.LBPHFaceRecognizer_create()
recognizer.read('trainer.yml')

label_id = {}
with open('labels.pickle', 'rb') as f:
	original_label_id = pickle.load(f)
	label_id = {a:b for b, a in original_label_id.items()}


cam = cv2.VideoCapture(0)
while True:
	ret, frame = cam.read()
	img = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

	#enables you to detect features
	faces = face_cascade.detectMultiScale(img)

	for x, y, w, h in faces:
		cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 3)
		roi = img[y:y+h, x:x+w]

		id_, conf = recognizer.predict(roi)
		if conf >= 45:
			font = cv2.FONT_HERSHEY_SIMPLEX

			name = label_id[id_]
			date = today.strftime("%d")

			colour = (255, 255, 255)
			cv2.putText(frame, name, (x, y), font, 1, colour, 2)

			sheet.cell(row = day_numbers[date], column = roll_numbers[name]).value = 1
			attendance_sheet.save('Attendance.xlsx')


	cv2.imshow("my camera", frame)
	if cv2.waitKey(1) == 13: 
		break
cam.release()
cv2.destroyAllWindows()