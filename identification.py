import os
import cv2
from PIL import Image
import numpy as np
import pickle
from sheet_generator import generate_sheet

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
image_dir = os.path.join(BASE_DIR, "images")

number_of_students = len(next(os.walk(image_dir))[1])
generate_sheet(number_of_students)

face_cascade = cv2.CascadeClassifier('face.xml')
recognizer = cv2.face.LBPHFaceRecognizer_create()

y_labels = []
x_train = []

label_id = {}
currentid = 0

for root, dirs, files in os.walk(image_dir):
	for file in files:
		if file.endswith("png") or file.endswith("jpg") or file.endswith("jpeg"):
			path = os.path.join(root, file)
			label = os.path.basename(root)

			if label not in label_id:
				label_id[label] = currentid
				currentid += 1

			#convert to gray
			pil_image = Image.open(path).convert("L")
			image_array = np.array(pil_image, "uint8")

			faces = face_cascade.detectMultiScale(image_array)
			for x, y, w, h in faces:
				roi = image_array[y:y+h, x:x+w]
				x_train.append(roi)
				y_labels.append(label_id[label])

with open("labels.pickle", "wb") as f:
	pickle.dump(label_id, f)

recognizer.train(x_train, np.array(y_labels))
recognizer.save('trainer.yml')
