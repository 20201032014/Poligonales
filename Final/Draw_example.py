from re import M
from turtle import Turtle, mainloop

my_pen = Turtle()
my_pen.pencolor("blue")
my_pen.hideturtle()
my_pen.penup()
my_pen.goto(100,50)
my_pen.showturtle()
my_pen.pendown()
my_pen.forward(200)
my_pen.left(90)
my_pen.forward(150)
my_pen.left(90)
my_pen.forward(200)
my_pen.left(90)
my_pen.forward(150)
my_pen.hideturtle() #esconde la flecha
mainloop() #mantiene la ventana abierta cuando termina