from tkinter import *
from PIL import ImageTk, Image

root = Tk()
root.title('Image Viewer')
root.iconbitmap('StoManager.ico')

my_img1 = ImageTk.PhotoImage(Image.open("rD:\STOMATA\Model training\New folder (2)\B1, R3, T7, Oneal, 1, 20X.jpg"))
my_img2 = ImageTk.PhotoImage(Image.open(r"D:\STOMATA\Model training\New folder (2)\B1, R3, T7, Oneal, 2, 20X.jpg"))
my_img3 = ImageTk.PhotoImage(Image.open(r"D:\STOMATA\Model training\New folder (2)\B1, R3, T7, Oneal, 3, 20X.jpg"))
my_img4 = ImageTk.PhotoImage(Image.open(r"D:\STOMATA\Model training\New folder (2)\B1, R3, T7, Oneal, 4, 20X.jpg"))
my_img5 = ImageTk.PhotoImage(Image.open(r"D:\STOMATA\Model training\New folder (2)\B1, R3, T7, Oneal, 5, 20X.jpg"))

image_list= [my_img1,my_img2,my_img3,my_img4,my_img5]


my_label = Label(image=my_img1)
my_label.grid(row=0,column=0,columnspan=3)

def forward(image_number):
    global my_label
    global button_forward
    global button_back

    my_label.grid_forget()
    my_label = Label(image_list[image_number-1])
    button_forward = Button(root,text=">>", command=lambda:forward(image_number+1))
    button_back = Button(root,text='<<',command=lambda:back(image_number-1))
    my_label.grid(row=0, column=0, columnspan=3)



def back():
    global my_label
    global button_forward
    global button_back





button_back=Button(root, text="<<",command=back())
button_exit = Button(root,text="EXIT PROGRAM",command=root.quit)
button_forward=Button(root, text=">>", command=lambda:forward(2))


button_back.grid(row=1,column=0)
button_exit.grid(row=1, column=1)
button_forward.grid(row=1,column=2)

root.mainloop()
