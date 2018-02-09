python

convert Excel to PDF
	## Author: Sirvan Almasi Jan 2017
	## This script helps in automating the process of converting an excel into PDF
	import win32com.client, time

	o = win32com.client.Dispatch("Excel.Application")
	o.Visible = False
	timedate = time.strftime("%H%M__%d_%m_%Y")
	wb_path = r'S:/GSA Euro Research Company Files/Property Sectors/Euro Office Sector/London Offices/Green Street Research Reports/London Office Report Feb 17/_5. Appendix - Company Snapshots - Copy.xlsm'
	#wb_path = r'C:/Users/salmasi/Documents/MATLAB/xlstopdf/22.xlsm'
	wb = o.Workbooks.Open(wb_path)

	ws_index_list = [1,2,3] #say you want to print these sheets
	path_to_pdf = r'C:/Users/salmasi/Documents/MATLAB/xlstopdf/app__'+str(timedate)+'.pdf'
	wb.WorkSheets(ws_index_list).Select()
	wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
	wb.Close(True)


copy entire excel worksheet to a new worksheet using Python win32com
	# old_sheet: sheet that you want to copy
	old_sheet.Copy(pythoncom.Empty, workbook.Sheets(workbook.Sheets.Count)
	new_sheet = workbook.Sheets(workbook.Sheets.Count)
	new_sheet.Name = 'Annual'	

Instead of using the PrintOut method, use ExportAsFixedFormat. You can specify the pdf format and supply a file name. Try this:
	ws.ExportAsFixedFormat(0, 'c:\users\alex\foo.pdf')	






Print chosen worksheets in excel files to pdf in python

	import win32com.client

	o = win32com.client.Dispatch("Excel.Application")

	o.Visible = False

	wb_path = r'c:\user\desktop\sample.xls'

	wb = o.Workbooks.Open(wb_path)



	ws_index_list = [1,4,5] #say you want to print these sheets

	path_to_pdf = r'C:\user\desktop\sample.pdf'



	wb.WorkSheets(ws_index_list).Select()

	wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)


	


Opencv and python for auto cropping

	If you want to do this with OpenCV, a good starting point may be after doing some simple processing to remove noise and small details in the image, you can find the edges of the image and then find the bounding box and crop to that area. But in case of your second image, you may need to do some post-processing as the raw edges may hold some noise and borders. You can do this on a pixel-by-pixel basis, or another maybe overkill method would be finding all the contours in the image and the finding the biggest bounding box. Using this you can get the following results: First Image
	And for the second one:
	Second Image
	The part that needs work is finding a proper thresholding method that works for all the images. Here I used different thresholds to make a binary image, as the first one was mostly white and second one was a bit darker. A first guess would be using the average intensity as a clue.
	Hope this helps!
	
	This is how I used some pre-processing and also a dynamic threshold to get it work for both of the images:

	im = cv2.imread('cloth.jpg')
	imgray = cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)
	imgray = cv2.blur(imgray,(15,15))
	ret,thresh = cv2.threshold(imgray,math.floor(numpy.average(imgray)),255,cv2.THRESH_BINARY_INV)
	dilated=cv2.morphologyEx(thresh, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_ELLIPSE,(10,10)))
	_,contours,_ = cv2.findContours(dilated,cv2.RETR_LIST,cv2.CHAIN_APPROX_SIMPLE)
	I also checked the contour area to remove very large contours:

	new_contours=[]
	for c in contours:
	    if cv2.contourArea(c)<4000000:
	        new_contours.append(c)
	The number 4000000 is an estimation of the image size (width*height), big contours should have an area close to the image size.

	Then you can iterate all the contours, and find the overall bounding box:

	best_box=[-1,-1,-1,-1]
	for c in new_contours:
	   x,y,w,h = cv2.boundingRect(c)
	   if best_box[0] < 0:
	       best_box=[x,y,x+w,y+h]
	   else:
	       if x<best_box[0]:
	           best_box[0]=x
	       if y<best_box[1]:
	           best_box[1]=y
	       if x+w>best_box[2]:
	           best_box[2]=x+w
	       if y+h>best_box[3]:
	           best_box[3]=y+h
	Then you have the bounding box of all contours inside the best_box array.
	https://stackoverflow.com/questions/37803903/opencv-and-python-for-auto-cropping








How to detect edge and crop an image in Python
	What you need is thresholding. In OpenCV you can accomplish this using cv2.threshold().

	I took a shot at it. My approach was the following:

	Convert to grayscale
	Threshold the image to only get the signature and nothing else
	Find where those pixels are that show up in the thresholded image
	Crop around that region in the original grayscale
	Create a new thresholded image from the crop that isn't as strict for display
	Here was my attempt, I think it worked pretty well.

	import cv2
	import numpy as np

	# load image
	img = cv2.imread('image.jpg') 
	rsz_img = cv2.resize(img, None, fx=0.25, fy=0.25) # resize since image is huge
	gray = cv2.cvtColor(rsz_img, cv2.COLOR_BGR2GRAY) # convert to grayscale

	# threshold to get just the signature
	retval, thresh_gray = cv2.threshold(gray, thresh=100, maxval=255, type=cv2.THRESH_BINARY)

	# find where the signature is and make a cropped region
	points = np.argwhere(thresh_gray==0) # find where the black pixels are
	points = np.fliplr(points) # store them in x,y coordinates instead of row,col indices
	x, y, w, h = cv2.boundingRect(points) # create a rectangle around those points
	x, y, w, h = x-10, y-10, w+20, h+20 # make the box a little bigger
	crop = gray[y:y+h, x:x+w] # create a cropped region of the gray image

	# get the thresholded crop
	retval, thresh_crop = cv2.threshold(crop, thresh=200, maxval=255, type=cv2.THRESH_BINARY)

	# display
	cv2.imshow("Cropped and thresholded image", thresh_crop) 
	cv2.waitKey(0)