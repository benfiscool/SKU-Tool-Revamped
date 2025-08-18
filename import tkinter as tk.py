from tkinter import *
import re

root=Tk()


def fun1(event, *args, **kwargs):
      global data
      if (e.get()==''):
            lb.place_forget()            
            lb.delete(0, END)
       
      else:
            lb.delete(0, END)
            value=e.get()
            lb.place(x=0, y=20)

            for items in data:
                        if (re.search(value, items, re.IGNORECASE)):    
                                    lb.insert(END, items)
                                
        
            print(value)
            pass
        
        



def CurSelet(evt):

  

    valued=lb.get(ACTIVE)
    e.delete(0, END)
    e.insert(END, valued)
    lb.place_forget() 

    print( valued)
        





def  down(ddd):
 
                  lb.focus()
                  lb.selection_set(0)

s=StringVar()


e=Entry(root,     textvariable=s)
e.grid(row=2, column=2)

s.trace('w', fun1)


e.bind('<Down>',     down)


for i in range(4,12):
      ee=Entry(root)
      ee.grid(row=i, column=2)



data=['Angular', 'action Script', 'Basic', 'GW-Basic' , 'C', 'C++', 'C#', 
'Django' ,'Dot-Net',  'Flask' , 'Go-Lang', 'Html', 'Python', 'PHP', 'Pearl', 
'Java', 'Javascript', 'Kotlin',  'Rust', 'R', 'S', 'Sr', 'Tekken 7', 'Tekken', 
'Tag' ]

lb=Listbox(root)
lb.place_forget()
lb.bind("<Button-3>", CurSelet)
lb.bind("<Right>",  CurSelet)


root.mainloop()

print(Listbox.curselection)