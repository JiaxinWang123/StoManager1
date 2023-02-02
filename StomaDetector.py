from tkinter import *
from PIL import ImageTk, Image

root = Tk()

root.title('StoManager')
root.iconbitmap('stomata_F0X_icon.ico')
root.geometry("1024x760")


bg = ImageTk.PhotoImage(file="D:/STOMATA/New folder (2)/T35 40x #8.jpg")

my_canvas=Canvas(root,width=1024,height=760,bd=0,highlightthickness=0)
my_canvas.pack(fill="both",expand=True)


my_canvas.create_image(0,0, image=bg, anchor="nw")

Input_path_entry=Entry(root,font=("Helvetica",24),width=28, fg="#06443B",bd=0)
Output_path_entry=Entry(root,font=("Helvetica",24),width=28, fg="#06443B",bd=0)

Input_path_entry.insert(0, "Type your image input folder path")
Output_path_entry.insert(0, "Type your file(s) output folder path")

Input_path_window = my_canvas.create_window(260,290,anchor="nw",window=Input_path_entry)
Output_path_window = my_canvas.create_window(260,370,anchor="nw",window=Output_path_entry)

def Input_entry_clear(e):
    if Input_path_entry.get() == "Type your image input folder path":
        Input_path_entry.delete(0, END)

def Output_entry_clear(e):
    if Input_path_entry.get() != "Type your image input folder path":
        Output_path_entry.delete(0, END)

def run_analyze():
    Input_path = Input_path_entry.get()
    Output_path = Output_path_entry.get()

    import cv2
    import numpy as np
    import glob
    import random
    import pandas as pd
    import math
    import os

    # Load Yolo
    net = cv2.dnn.readNet("yolov3_training_last (5).weights", "yolov3_testing.cfg")

    # Name custom object
    classes = ["whole_stomata", "stomata"]

    # Images path
    #image_path_input = input("Please type your image path:")
    image_path_input=Input_path
    images_path = glob.glob(str(image_path_input) + r"\*.jpg")

    print(images_path)
    layer_names = net.getLayerNames()
    output_layers = [layer_names[i - 1] for i in net.getUnconnectedOutLayers()]
    colors = np.random.uniform(0, 255, size=(len(classes), 3))
    # Insert here the path of your images
    random.shuffle(images_path)
    # loop through all the images

    save_files = []
    if bool(save_files) == False:
        #save_file = input('Please type the path where you want to save the export csv and jpg files:')
        save_file=Output_path
        save_files.append(save_file)
        save_file_str = ''.join(map(str, save_files))

    for img_path in images_path:
        # Loading image
        img = cv2.imread(img_path)
        # img = cv2.resize(img, None, fx=0.5, fy=0.5)
        height, width, channels = img.shape

        # Detecting objects
        blob = cv2.dnn.blobFromImage(img, 0.00392, (416, 416), (0, 0, 0), True, crop=False)

        net.setInput(blob)
        outs = net.forward(output_layers)
        # Showing informations on the screen
        class_ids = []
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
        for out in outs:
            for detection in out:
                scores = detection[5:]
                class_id = np.argmax(scores)
                confidence = scores[class_id]
                if confidence > 0.3:
                    # Object detected

                    center_x = int(detection[0] * width)
                    center_y = int(detection[1] * height)
                    w = int(detection[2] * width)
                    h = int(detection[3] * height)
                    # Rectangle coordinates
                    x = int(center_x - w / 2)
                    y = int(center_y - h / 2)

                    boxes.append([x, y, w, h])
                    confidences.append(float(confidence))
                    class_ids.append(class_id)

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
                cv2.rectangle(img, (x, y), (x + w, y + h), color, 1)
                cv2.putText(img, label, (x, y + 30), font, 0.8, color, 1)
                cv2.putText(img, str(h), (x, y + 50), font, 0.8, color, 1)
                cv2.putText(img, str(w), (x, y + 20), font, 0.8, color, 1)
                # cv2.putText(img, str(confidence), (x, y - 5), font, 1, color, 1)
                if label == "whole_stomata":
                    number_of_whole_stomata.append(class_id)
                elif label == "stomata":
                    number_of_stomata.append(class_id)
                orientation = math.degrees(np.arctan(h / w))
                orientations.append(orientation)
                list_of_width.append(w)
                list_of_height.append(h)
                image_paths.append(img_path)
                labels.append(label)
                list_of_image_height.append(height)
                list_of_image_width.append(width)
                list_of_all_stomata_areas.append(
                    (w * h + 604.69) / 1.3915)  ## Build linear regression model for adjusted area of stomata
        print("The current detecting image is :", img_path)
        print("The total number of images is:", len(images_path))
        print("There are: ", len(indexes), "objects in total.")
        # print("There are :", len(number_of_stomata), "stomatas.")
        # print("There are :", len(number_of_whole_stomata), "whole_stomatas.")
        print("There are :", len(number_of_stomata), "stomatas,", " and", len(number_of_whole_stomata),
              "whole_stomatas.", "in", img_path)
        # print("Currently detecting image is :",img_path)
        print("The measured size of this image is :", width, " × ", height, ".")
        heights = list_of_height
        widths = list_of_width
        image_name = img_path
        image_width = list_of_image_width
        image_height = list_of_image_height
        all_stomata_areas = list_of_all_stomata_areas
        image_name = {"Labels": labels, "Width": widths, "Height": heights, "Orientation": orientations,
                      "Num_of_Stomatas": len(number_of_stomata), "Num_of_whole_stomatas": len(number_of_whole_stomata),
                      "Width_of_image": image_width, "Height_of_image": image_height,
                      "Whole_sotmata_area": all_stomata_areas}

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

    cv2.destroyAllWindows()


Input_button= Button(root,text="Input Image Folder", font=("Helvetica",20), width=15, fg="#06443B")
Output_button= Button(root,text="Output Image Folder", font=("Helvetica",20), width=15, fg="#06443B")
Run_button=Button(root,text="Run!", font=("Helvetica",20), width=15, fg="#06443B",command= run_analyze)




Run_window= my_canvas.create_window(260,450,anchor="nw",window=Run_button,width=505)




Input_path_entry.bind("<Button-1>", Input_entry_clear)
Output_path_entry.bind("<Button-1>", Output_entry_clear)


root.mainloop()
