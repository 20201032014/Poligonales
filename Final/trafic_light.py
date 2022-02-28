from tkinter import Tk, Canvas
from tkinter.ttk import Button, Frame
from turtle import width
def do_button_press():
    global color

    if color== "red":
        color = "green"
        my_canvas.itemconfigure(red_lamp, fill="black")
        my_canvas.itemconfigure(green_lamp, fill="green")
    
    elif color == "green":
        color ="yellow"
        my_canvas.itemconfigure(green_lamp, fill="black")
        my_canvas.itemconfigure(yellow_lamp, fill="yellow")
    elif color == "yellow":
        color ="red"
        my_canvas.itemconfigure(red_lamp, fill="red")
        my_canvas.itemconfigure(yellow_lamp, fill="black")
color = "red"
my_form = Tk()
my_form.title("semaforo")
my_frame = Frame(my_form)
my_frame.pack()

my_canvas = Canvas(my_frame, width=150, height=300)

my_canvas.create_rectangle(50,21,150,280, fill="gray")
#my_canvas.create_rectangle(0,0,50,20, fill="blue")

red_lamp = my_canvas.create_oval(70,40,130,100, fill= "red")
yellow_lamp = my_canvas.create_oval(70,120,130,180, fill= "yellow")
green_lamp = my_canvas.create_oval(70,200,130,260, fill= "green")

my_button = Button(my_frame, text="Cambiar", command=do_button_press)

my_button.grid(row=1, column=0)
my_canvas.grid(row=0, column=0)
my_form.mainloop()