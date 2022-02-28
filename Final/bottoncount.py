from tkinter import Tk, Button

total_clicks = 0

def update():
    global total_clicks, my_button
    total_clicks += 1
    my_button.config(text="total clics = " + str(total_clicks))
    print("Actualizado")

  
my_form = Tk()
my_button = Button()
my_button.configure ( background="yellow", text="total clics = 0",command=update )
my_button.pack()
my_form.mainloop()