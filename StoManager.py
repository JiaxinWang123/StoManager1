from tkinter import *
from PIL import ImageTk, Image
import cv2
import numpy as np
import glob
import random
import pandas as pd
import math
import os
from tkinter import filedialog
import statistics
import string
from openpyxl.workbook import Workbook
import matplotlib.pyplot as plt
import re
from pandas import Series


# The main window


root = Tk()  ## Create a windows app using TK Moduel
root.resizable(width=False, height=False)
root.title('StoManager_V.02.10.23_for_Different_Trees')  # Title of the app
root.iconbitmap('StoManager.ico')  # the icon that will be showing on the topleft of the app
root.geometry("1024x760")  # the size of our windows app displayed on our screen
bg = ImageTk.PhotoImage(file="T35 40x #7.jpg")  # to set the background of our app

my_canvas = Canvas(root, width=1024, height=760, bd=0, highlightthickness=0)  # create  canvas in the windows of our app
my_canvas.pack(fill="both", expand=True)

# Create menus
# main_menu = Menu(my_canvas, tearoff=False)
# root.config(menu=main_menu)


# Create a frame
frame = LabelFrame(my_canvas, padx=5, pady=5, width=1024, height=760)
frame.pack(pady=7, padx=7)

my_canvas.create_image(0, 0, image=bg, anchor="nw")

Input_path_entry = Entry(root, font=("Helvetica", 24), width=28, fg="gray", bd=1)
Output_path_entry = Entry(root, font=("Helvetica", 24), width=28, fg="gray", bd=1)

Input_path_entry.insert(0, "Select an image input folder path")
Output_path_entry.insert(0, "Select an output folder path")


Input_img_size_entry = Entry(root, font=("Helvetica", 12), width=24, fg="black", bd=1)
Input_pixels_in_1_over_10_mm = Entry(root, font=("Helvetica", 12), width=24, fg="black", bd=1)

Input_img_size_entry.insert(0, "Input img full size: width, height")
Input_pixels_in_1_over_10_mm.insert(0, "Input pixels in 0.1 mm line")

Input_img_size_window = my_canvas.create_window(30, 30, anchor="nw",
                                            window=Input_img_size_entry)  # create the entry box in our canvas
Input_pixel_window = my_canvas.create_window(30, 55, anchor="nw", window=Input_pixels_in_1_over_10_mm)


confidence_entry = Entry(root, font=("Helvetica", 12), width=24, fg="black", bd=1)

confidence_entry.insert(0, "The default threshold is 0.3")

confidence_window = my_canvas.create_window(770, 30, anchor="nw",
                                            window=confidence_entry)  # create the entry box in our canvas




#### Create a label and then put it into windows of canvas
Label1 = Label(root, text="Copyright © Jiaxin Wang, email: coolwjx@foxmail.com", font=("Helvetica", 12), width=42,
               fg="gray", bd=0)

font = ("poppins", 15, "bold")
Label_info = Label(root, font=font)

Label_info.pack()
Label_window1 = my_canvas.create_window(330, 710, anchor="nw", window=Label1)
Label_show_info = my_canvas.create_window(375, 600, anchor="nw", window=Label_info)

Label_print = Label(root, font=font, fg="green")
Label_print.pack()
Label_print_info = my_canvas.create_window(880, 660, anchor="nw", window=Label_print)

Input_path_window = my_canvas.create_window(260, 290, anchor="nw",
                                            window=Input_path_entry)  # create the entry box in our canvas
Output_path_window = my_canvas.create_window(260, 370, anchor="nw", window=Output_path_entry)
Input_path = Input_path_entry.get()
Output_path = Output_path_entry.get()
Output_path = Output_path


# define a function to normalize the file names
def NormalizeFileNames():
    Input_path = Input_path_entry.get()
    
    Site = str()
    Block = str()
    Year = str()
    Month = str()
    Tail = str()
    Clone = str()
    file_names = str()

    # create list of Site, Block, Year, Month, Tail, and Clone here, or import those information from .txt files and read them as lists

    site = ["M","Mon","m","mon","P","p","Pon","pon","PON","M ","Mon ","m ","mon ","P ","p ","Pon ","pon ","PON "] # create a list containing all potential sites
    block = ["b1","b2","b5","B1","B2","B5","b1 ","b2 ","b5 ","B1 ","B2 ","B5 "," b1"," b2"," b5"," B1"," B2"," B5"] # create a list containing all potential blocks
    year = ["20","21","22","23","24"," 20"," 21"," 22"," 23"," 24","20 ","21 ","22 ","23 ","24 ","2022","2020","2021","2019"] # create a list containing all potential years
    month = ["June","june","July","july","Aug","aug","AUG","JULY","JUNE"," June"," june"," July"," july"," Aug"," aug",
    " AUG"," JULY"," JUNE","June ","june ","July ","july ","Aug ","aug ","AUG ","JULY ","JUNE ","Sept","Sep"," Sept","sep"," sep"] # create a list containing all potential months
    tail = ["1","2","3","4","5","6","7","8","9","10","Aci","Aci1","Aci2","Aci3","Aci4","Aci5","Aci6","Aci7","Aci8","Aci9",
    "Aci10","aci","aci1","aci2","aci3","aci4","aci5","aci6","aci7","aci8","aci9","aci10","ACi","aCI"," 1"," 2"," 3"," 4"," 5",
    " 6"," 7"," 8"," 9"," 10"," Aci"," Aci1"," Aci2"," Aci3"," Aci4"," Aci5"," Aci6"," Aci7"," Aci8"," Aci9"," Aci10"," aci",
    " aci1"," aci2"," aci3"," aci4"," aci5"," aci6"," aci7"," aci8"," aci9"," aci10"," ACi"," aCI","1 ","2 ","3 ","4 ","5 ","6 ",
    "7 ","8 ","9 ","10 ","Aci ","Aci1 ","Aci2 ","Aci3 ","Aci4 ","Aci5 ","Aci6 ","Aci7 ","Aci8 ","Aci9 ","Aci10 ","aci ","aci1 ",
    "aci2 ","aci3 ","aci4 ","aci5 ","aci6 ","aci7 ","aci8 ","aci9 ","aci10 ","ACi ","aCI ","40X","20X"," 40X"," 20x"," 40X ",
    " 20X ","40x ","20X ","EC","DN","DM","DD"] # create a list containing all potential tails
    clone = ["S7C2","s7c2","s7c4","S7C4","ST-66","st-66","st66","ST66","19","22","110412","11412","120-4","3-1","6-1","6-5","ST-70",
    "ST-75","st-70","st-75","st70","st75","6323","6329","8019","9225","9707","11690","13693","13724","24033","24056","24066","24114",
    "24120","24159","29310","433","11785","11789","11795","11797","11802","11840","13849","14278","14340","24250","7903","8729","10016",
    "24245","24301","24304","24326","24340","29262","29270","11666","9709","14492","11691","13700","14486","8015","8002","14507","13725",
    "11711","8852","10243","9225","14591","6323","13695","9711","14490","7388","9189","11732","11822","433","11867","11859","7389","8198",
    "8729","7903","10029","9552","9755","9980","8717","11733","111733","120-4","S7C20","s7c20","ST75","1BB-3","113B-3","113b-3","112107",
    "S13C20","s13c20","106B-1","106b-1","47-5","6-4","9198","7938"] # create a list containing all potential clones, you can also import clone.txt file and read it as a list

    for file_path in glob.glob(Input_path + "/" + '*jpg'):
        file_name = file_path[len(Input_path) + 1:-4]  # Extract the file name from the file path
        file_name_split = file_name.split(",")
        Split_name = []
        for s in file_name_split:
            s = re.sub("\s+", ",", s.strip())
            Split_name.append(s)
        Split_name = ",".join(Split_name)
        Split_name= Split_name.split(",")
        print(Split_name)  # Split the site, block, and clone info from the file name

        if len(Split_name) >= 5:
            for name in Split_name:

                if name in site:
                    Site = name

                elif name in block:
                    Block = name

                elif name in year:
                    Year = name

                elif name in month:
                    Month = name

                elif name in tail:
                    Tail = name

                elif name in clone:
                    Clone = name
        elif len(Split_name) <=4:

            file_names =",".join([str(item) for item in Split_name])


        else:
            for name in Split_name:
                if name in site:
                    Site = name
                elif name in block:
                    Block = name
                elif name in year:
                    Year = name
                elif name in month:
                    Month = name
                elif name in clone:
                    Clone = name
                

        # Site = Site.upper()
        # Block= Block.upper()
        # Clone = Clone.upper()
        # Month = Month.upper()
        # Tail = Tail.upper()
        # file_names = file_names.upper()

        if len(Split_name) > 4:
            New_file_name = Site + "," + Block + "," + Clone + "," + Month + "," + Year + "," + Tail + ".jpg"
            if not os.path.exists(Input_path + "/" + New_file_name):
                os.rename(file_path, Input_path + "/" + New_file_name)
        else:
            New_file_name = file_names + ".jpg"
            if not os.path.exists(Input_path + "/" + New_file_name):
                os.rename(file_path, Input_path + "/" + New_file_name)


        
       
            


def run_analyze():  ## the main funtion with the yolov3 model we trained to detect our stomata
    
    
    Input_path = Input_path_entry.get()
    Output_path = Output_path_entry.get()

    Input_img_size = Input_img_size_entry.get()
    Input_pixels = Input_pixels_in_1_over_10_mm.get()

    if Input_img_size =="Input img full size: width, height" or Input_img_size =="" or " ":
        Input_img_size = "2048,1536"
    else: 
        Input_img_size = Input_img_size_entry.get()  

    if Input_pixels =="Input pixels in 0.1 mm line" or Input_pixels =="" or " ":
        Input_pixels = "476"
    else: 
        Input_pixels = Input_pixels_in_1_over_10_mm.get()   

    split_img_size = Input_img_size.split(",")

    pixle = float(Input_pixels)

    Img_area = float(split_img_size[0])*float(split_img_size[1])
    
    whole_img_area = Img_area/((pixle)*10.0)**2

    # Load Yolo
    net = cv2.dnn.readNet("yolov3_training_last.weights", "yolov3_testing.cfg")

    # Name custom object
    classes = ["whole_stomata", "stomata"]

    # Images path
    # image_path_input = input("Please type your image path:")
    image_path_input = Input_path
    images_path = glob.glob(str(image_path_input) + r"\*.jpg")

    print(images_path)
    layer_names = net.getLayerNames()
    output_layers = [layer_names[i - 1] for i in net.getUnconnectedOutLayers()]
    colors = np.random.uniform(0, 100, size=(len(classes), 3))
    # Insert here the path of your images
    random.shuffle(images_path)
    # loop through all the images

    save_files = []
    if bool(save_files) == False:
        # save_file = input('Please type the path where you want to save the export csv and jpg files:')
        save_file = Output_path
        save_files.append(save_file)
        save_file_str = ''.join(map(str, save_files))

    for img_path in images_path:
        # Loading image
        confidence_set = confidence_entry.get()
    
        if confidence_set =="The default threshold is 0.3" or confidence_set =="" or " ":
            confidence_set = 0.3
        else: 
            confidence_set = float(confidence_entry.get())
        

        img = cv2.imread(img_path)
        # img = cv2.resize(img, None, fx=0.5, fy=0.5)
        height, width, channels = img.shape

        # Detecting objects
        blob = cv2.dnn.blobFromImage(img, 0.00392, (416, 416), (0, 0, 0), True, crop=False)

        net.setInput(blob)
        outs = net.forward(output_layers)
        # Showing information on the screen
        class_ids = []
        class_ids_2 = []
        centers = []
        confidences = []
        boxes = []
        number_of_whole_stomata = []
        number_of_stomata = []
        list_of_width = []
        list_of_height = []
        image_paths = []
        orientations = []
        labels = []
        list_of_image_width = []
        list_of_image_height = []
        list_of_all_stomata_areas = []
        list_of_whole_stomatal_area = []
        list_of_x = []
        list_of_y = []
        Whole_stomata_density = []
        for out in outs:
            for detection in out:
                scores = detection[5:]
                class_id = np.argmax(scores)
                confidence = scores[class_id]
                if confidence > confidence_set:
                    # Object detected

                    center_x = (detection[0] * width)
                    center_y = (detection[1] * height)
                    w = (detection[2] * width)
                    h = (detection[3] * height)
                    # Rectangle coordinates
                    x = (center_x - w / 2)
                    y = (center_y - h / 2)

                    boxes.append([x, y, w, h])
                    centers.append([center_x, center_y])
                    confidences.append(float(confidence))
                    class_ids.append(class_id)
        # put class id, confidence, x, y, width, height to a txt file
        
        print(class_ids)
        # print(confidences)
        # print(scores)
        indexes = cv2.dnn.NMSBoxes(boxes, confidences, 0.5, 0.4)  

        # print(len(indexes[class_id]))
        font = cv2.FONT_HERSHEY_PLAIN
        for i in range(len(boxes)):

            if i in indexes:
                x, y, w, h = boxes[i]
                label = str(classes[class_ids[i]])
                color = colors[class_ids[i]]
                
                cv2.rectangle(img, (int(x), int(y)), (int(x) + int(w), int(y) + int(h)), color, 1)
                cv2.putText(img, label, (int(x), int(y) + 42), font, 1, color, 1)
                cv2.putText(img, str(int(h)), (int(x), int(y) + 60), font, 1, color, 1)
                cv2.putText(img, str(int(w)), (int(x), int(y) + 28), font, 1, color, 1)
                cv2.putText(img, str(round(confidences[i], 2)), (int(x), int(y) + 12), font, 1, color, 1)

                # cv2.putText(img, str(confidence), (x, y - 5), font, 1, color, 1)
                if label == "whole_stomata":
                    number_of_whole_stomata.append(class_ids[i])
                    list_of_all_stomata_areas.append(
                        (w * h) * 0.6878 + 806)  ## Build linear regression model for adjusted area of whole_stomata
                    list_of_whole_stomatal_area.append((w * h) * 0.6878 + 806)
                elif label == "stomata":  ## Build linear regression model for adjusted area of stomata
                    number_of_stomata.append(class_ids[i])
                    list_of_all_stomata_areas.append(
                        (w * h + 116.08) / 1.7684)
                    
                
                orientation = math.log(w/h)*(-92.2325)+44.5222
                orientations.append(orientation)

                list_of_width.append(w)
                list_of_height.append(h)
                list_of_x.append(x)
                list_of_y.append(y)
                image_paths.append(img_path)
                labels.append(label)
                list_of_image_height.append(height)
                list_of_image_width.append(width)
                class_ids_2.append(class_ids[i])
        
        
        print(labels)
        with open(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".txt", 'w') as f:
            
            for m in range(len(list_of_width)):
                w_2 = list_of_width[m]/width
                h_2 = list_of_height[m]/height
                x_center = (list_of_x[m]+list_of_width[m]/2)/width
                y_center = (list_of_y[m]+list_of_height[m]/2)/height


                x = (center_x - w / 2)

                f.write(str((str(class_ids_2[m]) + " " + str(confidences[m])+ " "+ str(x_center) + " " + str(y_center) + " " + str(w_2) + " " + str(h_2))))
                f.write('\n')  

        print("The current detecting image is :", img_path)
        print("The total number of images is:", len(images_path))
        print("There are: ", len(indexes), "objects in total.")
        # print("There are :", len(number_of_stomata), "stomatas.")
        # print("There are :", len(number_of_whole_stomata), "whole_stomatas.")
        print("There are :", len(number_of_stomata), "stomatas,", " and", len(number_of_whole_stomata),
              "whole_stomatas.", "in", img_path)
        # print("Currently detecting image is :",img_path)
        print("The measured size of this image is :", width, "x", height, ".")
        if len(images_path) > 1:
            final_info = "There are " + str(
                len(images_path)) + " images detected!" + "\n" + "It's time to check your results!"
            Label_info.config(text=final_info)
        else:
            final_info = " There is  " + str(
                len(images_path)) + "  image detected!" + "\n" + "It's time to check your result!"
            Label_info.config(text=final_info)

        list_of_whole_stomatal_area_ratio = sum(list_of_whole_stomatal_area) / (
                statistics.mean(list_of_image_width) * statistics.mean(list_of_image_height))
        heights = list_of_height
        widths = list_of_width
        image_name = img_path
        image_width = list_of_image_width
        image_height = list_of_image_height
        all_stomata_areas = [(float(k)/Img_area*1000000) for k in list_of_all_stomata_areas]
        Whole_stomata_density=number_of_whole_stomata
        image_name = {"Labels": labels, "Width_(pixels)": widths, "Height_(pixels)": heights,
                      "Orientation": orientations,
                      "Num_of_Stomata": len(number_of_stomata), "Num_of_whole_stomata": len(number_of_whole_stomata),
                      "Width_of_image_(pixels)": image_width, "Height_of_image_(pixels)": image_height,
                      "All_sotmata_area_(mum2)": all_stomata_areas,
                      "Whole_stomata_area_ratio": list_of_whole_stomatal_area_ratio,
                      "Whole_stomata_density": len(Whole_stomata_density)/whole_img_area}

        df = pd.DataFrame(image_name)

        # print(image_name)
        # df.to_csv('width_height.csv')
        img_path = str(img_path)
        new_img_path = img_path.replace("\\", "")
        new_img_path = img_path.replace(".jpg", "")
        print(new_img_path)

        file = df.to_csv(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".csv", sep=',', index=True)
        cv2.imwrite(str(save_file_str) + "\\" + os.path.basename(img_path)[:-4] + ".jpg", img)

        print(file)
        print(save_file)
        print(range(len(save_files)))
        print(bool(save_files))
        print(save_file_str)
        # cv2.imshow("Image", img)

        # key = cv2.waitKey(0)

    # cv2.destroyAllWindows()



# Define a function to analyze the  stomatal data

def Stomata():
    # Create empty lists to hold the values that we are going to extract
    folder_path = Output_path_entry.get()
    WST_Area_median = []
    WST_Area_max = []
    WST_Area_min = []
    WST_Area_mean = []
    WST_Area_var = []
    WST_Area_std = []
    WST_Number = []
    WST_Area_Ratio = []
    WST_Orientation = []
    WST_Orientation_var = []
    WST_Density = []
    WST_Area = []
    Site = []
    Block = []
    Clone = []
    Month = []
    Year = []

    # Using for loop to go through all csv files
    
    for file_path in glob.glob(folder_path + '/' + '*.csv'):
        single_csv_file = pd.read_csv(file_path, low_memory=False)  # Read the csv and assign it to a variable
        file_name = file_path[len(folder_path) + 1:-4]  # Extract the file name from the file path
        Split_name = file_name.split(",")  # Split the site, block, and clone info from the file name
        
        Site.append(Split_name[0])  # Add the site to the Site list
        Block.append(Split_name[1])
        Clone.append(Split_name[2])
        Month.append(Split_name[3])
        Year.append(Split_name[4])

        # Extract the parameters from each single csv files

        Orientation = single_csv_file["Orientation"]  # Extract the Orientation
        Whole_stomata_number = single_csv_file["Num_of_whole_stomata"]  # Extract the number of whole stomata
        Whole_stomata_areas = single_csv_file[single_csv_file['Labels'] == 'whole_stomata']["All_sotmata_area_(mum2)"]
        Whole_stomata_area_ratio = single_csv_file['Whole_stomata_area_ratio']
        Whole_stomata_density = single_csv_file["Whole_stomata_density"]

        # Remove the outliers before calculating orientation variance
        Q1 = np.percentile(Orientation, 25, method = 'midpoint')
        Q3 = np.percentile(Orientation, 75,method = 'midpoint')

        IQR = Q3 - Q1

        upper = np.where(Orientation >= (Q3+1.5*IQR))
        lower = np.where(Orientation <= (Q1-1.5*IQR))

        Orientation.drop(upper[0], inplace = True)
        Orientation.drop(lower[0], inplace = True)

        # print(upper)
        # print(lower)

        # Remove the outliers before calculating Whole_stomata_areas variance
        # Q1 = np.percentile(Whole_stomata_areas, 25, method = 'midpoint')
        # Q3 = np.percentile(Whole_stomata_areas, 75,method = 'midpoint')

        # IQR = Q3 - Q1

        # upper = np.where(Whole_stomata_areas >= (Q3+1.5*IQR))
        # lower = np.where(Whole_stomata_areas <= (Q1-1.5*IQR))

        # Whole_stomata_areas.drop(upper[0], inplace = True)
        # Whole_stomata_areas.drop(lower[0], inplace = True)

        # print(Orientation)
        # print(Whole_stomata_areas)


        # calculate the median, mean, variance of each leaf stomata number, stomata area

        Whole_stomata_areas_median = np.median(Whole_stomata_areas)
        Whole_stomata_areas_max = np.max(Whole_stomata_areas)
        Whole_stomata_areas_min = np.min(Whole_stomata_areas)
        Whole_stomata_areas_mean = np.mean(Whole_stomata_areas)
        Whole_stomata_areas_var = np.var(Whole_stomata_areas)
        Whole_stomata_areas_std = np.std(Whole_stomata_areas)

        WST_Area_median.append(Whole_stomata_areas_median)
        WST_Area_max.append(Whole_stomata_areas_max)
        WST_Area_min.append(Whole_stomata_areas_min)
        WST_Area_mean.append(Whole_stomata_areas_mean)
        WST_Area_var.append(Whole_stomata_areas_var)
        WST_Area_std.append(Whole_stomata_areas_std)
        WST_Orientation.append(np.median(Orientation))
        WST_Orientation_var.append(np.var(Orientation))

        # Extract the number of stomata
        WST_Density.append(np.mean(Whole_stomata_density))
        WST_Number.append(np.mean(Whole_stomata_number))
        WST_Area.append(np.mean(Whole_stomata_areas))
        WST_Area_Ratio.append(np.mean(Whole_stomata_area_ratio))

    # Convert all lists into pd.series

    Site = pd.Series(Site, dtype=pd.StringDtype(), name="Site")
    Block = pd.Series(Block, dtype=pd.StringDtype(), name="Block")
    Clone = pd.Series(Clone, dtype=pd.StringDtype(), name="Clone")
    Month = pd.Series(Month, dtype=pd.StringDtype(), name="Month")
    Year = pd.Series(Year, dtype=pd.StringDtype(), name="Year")
    WST_Number = pd.Series(WST_Number, dtype=pd.Float64Dtype(), name="WST_Number")
    WST_Area_Ratio = pd.Series(WST_Area_Ratio, dtype=pd.Float64Dtype(), name="WST_Area_Ratio")
    WST_Area_median = pd.Series(WST_Area_median, dtype=pd.Float64Dtype(), name="WST_Area_median")
    WST_Area_max = pd.Series(WST_Area_max, dtype=pd.Float64Dtype(), name="WST_Area_max")
    WST_Area_min = pd.Series(WST_Area_min, dtype=pd.Float64Dtype(), name="WST_Area_min")
    WST_Area_mean = pd.Series(WST_Area_mean, dtype=pd.Float64Dtype(), name="WST_Area_mean")
    WST_Area_var = pd.Series(WST_Area_var, dtype=pd.Float64Dtype(), name="WST_Area_var")
    WST_Area_std = pd.Series(WST_Area_std, dtype=pd.Float64Dtype(), name="WST_Area_std")
    WST_Density = pd.Series(WST_Density, dtype=pd.Float64Dtype(), name ="WST_Density")
    WST_Orientation = pd.Series(WST_Orientation, dtype=pd.Float64Dtype(), name ="WST_Orientation")
    WST_Orientation_var = pd.Series(WST_Orientation_var, dtype=pd.Float64Dtype(), name ="WST_Orientation_var")

    # Put all extracted parameters into a data frame
    Output_data = pd.concat(
        [Site, Block, Clone, Month, Year, WST_Number, WST_Area_Ratio,WST_Density, WST_Area_median, WST_Area_max, WST_Area_min,
         WST_Area_mean, WST_Area_var, WST_Area_std, WST_Orientation, WST_Orientation_var], axis=1)
    Output_data = pd.DataFrame(data=Output_data)
    random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
    Output_data.to_excel(folder_path + "/" + "Stomata_output"+ "_" + random_str +".xlsx")
    print_text = "Done!"
    Label_print.config(text=print_text)


def Stomata_with_no_groups():
    # Create empty lists to hold the values that we are going to extract
    folder_path = Output_path_entry.get()
    WST_Area_median = []
    WST_Area_max = []
    WST_Area_min = []
    WST_Area_mean = []
    WST_Area_var = []
    WST_Area_std = []
    WST_Number = []
    WST_Area_Ratio = []
    WST_Orientation = []
    WST_Orientation_var = []
    WST_Density = []
    WST_Area = []
    Filename = []


    # Using for loop to go through all csv files
    
    for file_path in glob.glob(folder_path + '/' + '*.csv'):
        single_csv_file = pd.read_csv(file_path, low_memory=False)  # Read the csv and assign it to a variable
        file_name = file_path[len(folder_path) + 1:-4]  # Extract the file name from the file path
        Split_name = file_name.split(",")  # Split the site, block, and clone info from the file name
        
        Filename.append(",".join([str(item) for item in Split_name]))

        # Extract the parameters from each single csv files

        Orientation = single_csv_file["Orientation"]  # Extract the Orientation
        Whole_stomata_number = single_csv_file["Num_of_whole_stomata"]  # Extract the number of whole stomata
        Whole_stomata_areas = single_csv_file[single_csv_file['Labels'] == 'whole_stomata']["All_sotmata_area_(mum2)"]
        Whole_stomata_area_ratio = single_csv_file['Whole_stomata_area_ratio']
        Whole_stomata_density = single_csv_file["Whole_stomata_density"]

        # Remove the outliers before calculating orientation variance
        Q1 = np.percentile(Orientation, 25, method = 'midpoint')
        Q3 = np.percentile(Orientation, 75,method = 'midpoint')

        IQR = Q3 - Q1

        upper = np.where(Orientation >= (Q3+1.5*IQR))
        lower = np.where(Orientation <= (Q1-1.5*IQR))

        Orientation.drop(upper[0], inplace = True)
        Orientation.drop(lower[0], inplace = True)

        # print(upper)
        # print(lower)

        # Remove the outliers before calculating Whole_stomata_areas variance
        # Q1 = np.percentile(Whole_stomata_areas, 25, method = 'midpoint')
        # Q3 = np.percentile(Whole_stomata_areas, 75,method = 'midpoint')

        # IQR = Q3 - Q1

        # upper = np.where(Whole_stomata_areas >= (Q3+1.5*IQR))
        # lower = np.where(Whole_stomata_areas <= (Q1-1.5*IQR))

        # Whole_stomata_areas.drop(upper[0], inplace = True)
        # Whole_stomata_areas.drop(lower[0], inplace = True)

        # print(Orientation)
        # print(Whole_stomata_areas)


        # calculate the median, mean, variance of each leaf stomata number, stomata area

        Whole_stomata_areas_median = np.median(Whole_stomata_areas)
        Whole_stomata_areas_max = np.max(Whole_stomata_areas)
        Whole_stomata_areas_min = np.min(Whole_stomata_areas)
        Whole_stomata_areas_mean = np.mean(Whole_stomata_areas)
        Whole_stomata_areas_var = np.var(Whole_stomata_areas)
        Whole_stomata_areas_std = np.std(Whole_stomata_areas)

        WST_Area_median.append(Whole_stomata_areas_median)
        WST_Area_max.append(Whole_stomata_areas_max)
        WST_Area_min.append(Whole_stomata_areas_min)
        WST_Area_mean.append(Whole_stomata_areas_mean)
        WST_Area_var.append(Whole_stomata_areas_var)
        WST_Area_std.append(Whole_stomata_areas_std)
        WST_Orientation.append(np.median(Orientation))
        WST_Orientation_var.append(np.var(Orientation))

        # Extract the number of stomata
        WST_Density.append(np.mean(Whole_stomata_density))
        WST_Number.append(np.mean(Whole_stomata_number))
        WST_Area.append(np.mean(Whole_stomata_areas))
        WST_Area_Ratio.append(np.mean(Whole_stomata_area_ratio))

    # Convert all lists into pd.series


    Filename = pd.Series(Filename, dtype=pd.StringDtype(), name="Filename")
    WST_Number = pd.Series(WST_Number, dtype=pd.Float64Dtype(), name="WST_Number")
    WST_Area_Ratio = pd.Series(WST_Area_Ratio, dtype=pd.Float64Dtype(), name="WST_Area_Ratio")
    WST_Area_median = pd.Series(WST_Area_median, dtype=pd.Float64Dtype(), name="WST_Area_median")
    WST_Area_max = pd.Series(WST_Area_max, dtype=pd.Float64Dtype(), name="WST_Area_max")
    WST_Area_min = pd.Series(WST_Area_min, dtype=pd.Float64Dtype(), name="WST_Area_min")
    WST_Area_mean = pd.Series(WST_Area_mean, dtype=pd.Float64Dtype(), name="WST_Area_mean")
    WST_Area_var = pd.Series(WST_Area_var, dtype=pd.Float64Dtype(), name="WST_Area_var")
    WST_Area_std = pd.Series(WST_Area_std, dtype=pd.Float64Dtype(), name="WST_Area_std")
    WST_Density = pd.Series(WST_Density, dtype=pd.Float64Dtype(), name ="WST_Density")
    WST_Orientation = pd.Series(WST_Orientation, dtype=pd.Float64Dtype(), name ="WST_Orientation")
    WST_Orientation_var = pd.Series(WST_Orientation_var, dtype=pd.Float64Dtype(), name ="WST_Orientation_var")

    # Put all extracted parameters into a data frame
    Output_data = pd.concat(
        [Filename, WST_Number, WST_Area_Ratio,WST_Density, WST_Area_median, WST_Area_max, WST_Area_min,
         WST_Area_mean, WST_Area_var, WST_Area_std, WST_Orientation, WST_Orientation_var], axis=1)
    Output_data = pd.DataFrame(data=Output_data)
    random_str = ''.join(random.choices(string.ascii_uppercase, k=4))
    Output_data.to_excel(folder_path + "/" + "Stomata_output"+ "_" + random_str +".xlsx")
    print_text = "Done!"
    Label_print.config(text=print_text)



#### Put a button to aqure the filedialog

def loadInputFolder():
    Inputfilefolder = filedialog.askdirectory()
    Input_path_entry.delete(0, END)
    Input_path_entry.insert(END, Inputfilefolder)


def loadOutputFolder():
    Outputfilefolder = filedialog.askdirectory()
    Output_path_entry.delete(0, END)
    Output_path_folder = Output_path_entry.insert(END, Outputfilefolder)


#### Check the output
def OpenOutputFolder():
    folder_path = Output_path_entry.get()
    OpenOutputfilefolder = filedialog.askopenfilename(initialdir=folder_path,
                                                      filetypes=(("Jpg files", "*.jpg"), ("All files", "*.*")))
    # show_image=PhotoImage(imag=OpenOutputfilefolder)


photo = PhotoImage(file=r"filefolder.png")
button1 = Button(root, text="Select folder", image=photo, command=loadInputFolder, height=36, width=40,
                 font=("Helvetica", 18), fg="#06443B", bd=0)
button1_window = my_canvas.create_window(770, 290, anchor="nw", window=button1)

button2 = Button(root, text="Select folder", image=photo, command=loadOutputFolder, height=36, width=40,
                 font=("Helvetica", 18), fg="#06443B", bd=0)
button1_window = my_canvas.create_window(770, 370, anchor="nw", window=button2)

Button_3 = Button(root, text="Start Stomatal Analysis with groups", font=("Helvetica", 12), width=30, height=1,
                  fg="green", command=Stomata)
Button_3_window = my_canvas.create_window(530, 655, anchor="nw", window=Button_3)

Button_4 = Button(root, text="Start Stomatal Analysis without groups", font=("Helvetica", 12), width=32,height= 1,
                  fg="green", command=Stomata_with_no_groups)
Button_4_window = my_canvas.create_window(220, 655, anchor="nw", window=Button_4)

LOGO = PhotoImage(file=r"StoManager_8080.png", )
img = LOGO.subsample(12)
Label3 = Label(root, image=img)
Label3_window = my_canvas.create_window(420, 20, anchor="nw", window=Label3)


def openfilename():
    # open file dialog box to select image
    # The dialogue box has a title "Open"
    filename = filedialog.askopenfilename(title='Open')
    return filename


# function to open a new window
# on a button click
def openNewWindow():
    # Toplevel object which will
    # be treated as a new window
    # Create a Label in New window
    newWindow = Toplevel(root)
    newWindow.iconbitmap('Stomanager.ico')  # the icon that will be showing on the topleft of the app
    frame = Frame(newWindow, width=600, height=400)
    frame.pack()
    img2 = ImageTk.PhotoImage(Image.open('StoManager_8080.png'))
    # Create a button and place it into the window using grid layout

    # A Label widget to show in toplevel
    # Open Image in Output folder
    # folder_path = Output_path_entry.get()
    # filename =filedialog.askopenfilename(filetypes=(("jpg files","*.jpg"),("csv files","*.csv"),("all files","*.*")))
    global img3
    f_types = [('all files', '*.*'), ('jpg files', '*.jpg'), ('csv files', '*.csv')]
    filename3 = filedialog.askopenfilename(filetypes=f_types)
    img3 = ImageTk.PhotoImage(file=filename3)
    # resize_image=img3.resize(1024,760), Image.ANTIALIAS)
    Label_4 = Label(image=img3)  # using Label
    label2 = Label(frame, image=img3)
    label2.pack()

    # Create csv file viewer
    # Create and set the GUI for the passScreen of the Password Manager.

    # sets the title of the
    # Toplevel widget
    newWindow.title("StoManager_View_Output")

    # sets the geometry of toplevel
    newWindow.geometry("1024x760")


button3 = Button(root, text="Check the output", command=openNewWindow, font=("Helvetica", 12), fg="green", bd=1)
button3_window = my_canvas.create_window(515, 530, window=button3)


# a button widget which will open a


def Input_entry_clear(e):  # the function that will get the text from the entry box
    if Input_path_entry.get() != " ":
        Input_path_entry.delete(0, END)


def Output_entry_clear(e):
    if Output_path_entry.get() != " ":
        Output_path_entry.delete(0, END)

def Img_size_entry_clear(e):
    if Input_img_size_entry.get() != " ":
        Input_img_size_entry.delete(0, END)

def Pixel_entry_clear(e):
    if Input_pixels_in_1_over_10_mm.get() != " ":
        Input_pixels_in_1_over_10_mm.delete(0, END)

def confidence_entry_clear(e):
    if confidence_entry.get() != " ":
        confidence_entry.delete(0, END)


Input_button = Button(root, text="Input Image Folder", font=("Helvetica", 20), width=15, fg="#06443B")
Output_button = Button(root, text="Output Image Folder", font=("Helvetica", 20), width=15, fg="#06443B")
Run_button = Button(root, text="Run!", font=("Helvetica", 20), width=15, fg="#06443B",
                    command=lambda: [NormalizeFileNames(), run_analyze()])

#### Creat a Frame
# frame1=Frame(my_canvas,width=100,highlightbackground='red',highlightthicknes=3)
# frame1.grid(row=0,column=0,padx=20,pady=20,ipadx=20,ipady=20)
# l1=Label(frame1,text='Image Processed',fg='blue',font=(16))
# l1.grid(row=0,column=0,padx=10,pady=10)
# textbox=Entry(frame1,fg='blue',font=(16))
# textbox.grid(row=0,column=1)
# button1=Button(frame1,text='Show Image',fg='blue',font=(16),command= show_image)
# button1.grid(row=1,column=1,sticky=W)


Run_window = my_canvas.create_window(260, 450, anchor="nw", window=Run_button, width=510)

Input_path_entry.bind("<Button-1>", Input_entry_clear)
Output_path_entry.bind("<Button-1>", Output_entry_clear)


Input_img_size_entry.bind("<Button-1>", Img_size_entry_clear)
Input_pixels_in_1_over_10_mm.bind("<Button-1>", Pixel_entry_clear)

confidence_entry.bind("<Button-1>", confidence_entry_clear)

root.mainloop()
