import cv2
import imutils
from PIL import Image
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

cam = cv2.VideoCapture(0)
 
while True:
	check, frame = cam.read()
	gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
	gray = cv2.GaussianBlur(gray, (5,5), 0)
	edged  = cv2.Canny(gray, 75,200)
	cnts = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
	cnts = imutils.grab_contours(cnts)
	cnts = sorted(cnts, key = cv2.contourArea, reverse = True)[:5]
	for c in cnts:
		peri = cv2.arcLength(c,True)
		approx = cv2.approxPolyDP(c, 0.02 * peri, True)
		if len(approx) == 4:
			screenCnt = approx
			cv2.drawContours(frame, [screenCnt], -2, (0,255,0), 2)
			x,y,w,h = cv2.boundingRect(c)
			ROI = edged[y:y+h,x:x+w]
			image_pil = Image.fromarray(ROI)
			if image_pil:
				try:
					print(pytesseract.image_to_string(image_pil))
				except:
					print("Cannot find text")


	cv2.imshow("Found Document",ROI)
	cv2.imshow("Live - Outline", frame)

	if cv2.waitKey(1) == ord('q'):
		break


cam.release()
cv2.destroyAllWindows()
