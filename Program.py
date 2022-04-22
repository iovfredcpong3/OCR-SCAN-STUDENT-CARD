from tkinter import *
from tkinter import filedialog, messagebox, ttk
from PIL import ImageTk,Image
import cv2
import pytesseract
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
import pandas as pd
import numpy as np
import tkinter as tk

#window
root = Tk()
root.title("Program Scan Student RSU CARD.")
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0)
root.geometry("500x650+700+300")

filename = ""

# gray
def gray(img):
    
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    return img

# blur
def blur(img) :
    img_blur = cv2.GaussianBlur(img,(5,5),0)
    return img_blur

# threshold
def threshold(img):
    #pixels with value below 100 are turned black (0) and those with higher value are turned white (255)
    img = cv2.threshold(img,0, 255,cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]    
    return img

#distanceTransform
def dis(img):
    dist = cv2.distanceTransform(img, cv2.DIST_L2, 5)
    dist = cv2.normalize(dist, dist, 0, 1.0, cv2.NORM_MINMAX)
    dist = (dist * 255).astype("uint8")
    dist = cv2.threshold(dist, 0, 255,cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    return dist

def contours_text(orig, contours):
    for cnt in contours: 
        x, y, w, h = cv2.boundingRect(cnt) 

        # Drawing a rectangle on copied image 
        rect = cv2.rectangle(orig, (x, y), (x + w, y + h), (0, 255, 255), 2) 
        

        # Cropping the text block for giving input to OCR 
        cropped = orig[y:y + h, x:x + w] 

        # Apply OCR on the cropped image 
        config = ('-l tha+eng --oem 1 --psm 3')
        text = pytesseract.image_to_string(cropped, config=config) 

        return text

def scanner():
    # configurations
    config = ('-l tha+eng --oem 1 --psm 6')
    # pytessercat
    text = pytesseract.image_to_string(im, config=config,)
    # print text
    text = text.split('\n')
    text
    
    im_gray = gray(im)
    im_blur = blur(im_gray)
    im_thresh = threshold(im_blur)
    im_dis = dis(im_thresh)

    contours, _ = cv2.findContours(im_thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE) 
    text = contours_text(im,contours)
    text = text.split("\n")
    info = []
    for ele in text:
        if ele.strip():
            info.append(ele)
            
    i = 0     
    len_info = len(info)
    for x in info:
        if len(x) == 7 :
            break
        else :
            i = i+1
        
        
        
    
    
    IDs = info[i]
    name = info[i+1]
    lname = info[i+2]
    area = info[i+3]
    bid = info[i+4]
    
    def saveedit():
        infoex = [[e1.get(),e2.get(),e3.get(),e4.get(),e5.get()]]
        cols = ['Student ID','Name','Last name','Banking ID','area']
        df =pd.read_excel('C:\\Users\\AAA\\Desktop\\pj ocr\\info.xlsx',engine="openpyxl")
        newdf = pd.DataFrame(data=infoex,columns=cols)
        frames = [df,newdf]
        con = pd.concat([df,newdf],ignore_index=True)
        writer = pd.ExcelWriter('C:\\Users\\AAA\\Desktop\\pj ocr\\info.xlsx', engine='xlsxwriter')
        con.to_excel(writer,index = False)
        writer.save()
        
        textl1.destroy()
        textl2.destroy()
        textl3.destroy()
        textl4.destroy()
        textl5.destroy()
        my_label.destroy()
        text2.destroy()
        btn_save.destroy()
        btn_edit.destroy()
        btn_clear.destroy()
        ed.destroy()
        
        
    
    def savetoexcel():
        global flebel
        infoex = [[IDs,name,lname,bid,area]]
        cols = ['Student ID','Name','Last name','Banking ID','area']
        df =pd.read_excel('C:\\Users\\AAA\\Desktop\\pj ocr\\info.xlsx',engine="openpyxl")
        newdf = pd.DataFrame(data=infoex,columns=cols)
        frames = [df,newdf]
        con = pd.concat([df,newdf],ignore_index=True)
        writer = pd.ExcelWriter('C:\\Users\\AAA\\Desktop\\pj ocr\\info.xlsx', engine='xlsxwriter')
        con.to_excel(writer,index = False)
        flebel=Label(root,text="Save Finished",fg="green",font=100)
        flebel.pack()
        writer.save()
        textl1.destroy()
        textl2.destroy()
        textl3.destroy()
        textl4.destroy()
        textl5.destroy()
        my_label.destroy()
        text2.destroy()
        btn_save.destroy()
        btn_edit.destroy()
        btn_clear.destroy()
        
        
        


    def edit_info():
        global e1,e2,e3,e4,e5,ed
        ed = Tk()

        ed.geometry("250x250") # set the root dimensions
        ed.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
        ed.resizable(0, 0) # makes the root window fixed in size.
        Label(ed,text="รหัสนักศึกษา",font=40).grid(row=0)
        Label(ed,text="ชื่อ",font=40).grid(row=1)
        Label(ed,text="นามสกุล",font=40).grid(row=2)
        Label(ed,text="เลขบัญชีธนาคาร",font=40).grid(row=3)
        Label(ed,text="สาขา",font=40).grid(row=4)
        
        e1=tk.Entry(ed)
        e1.insert(10, IDs)
        
        e2=tk.Entry(ed)
        e2.insert(10, name)
        
        e3=tk.Entry(ed)
        e3.insert(10, lname)
        
        e4=tk.Entry(ed)
        e4.insert(10, bid)
        
        e5=tk.Entry(ed)
        e5.insert(10, area)
        
        e1.grid(row=0, column=1)
        e2.grid(row=1, column=1)
        e3.grid(row=2, column=1)
        e4.grid(row=3, column=1)
        e5.grid(row=4, column=1)
        
        
        
        btn_saveed = Button(ed,text="Save Edit",command=saveedit).grid(row=5,column=1)
        
    global textl1,textl2,textl3,textl4,textl5
    textl1=Label(root,text="รหัสนักศึกษา : "+IDs,fg="green",font=30)
    textl1.pack()
    textl2=Label(root,text="ชื่อ : "+name,fg="green",font=30)
    textl2.pack()
    textl3=Label(root,text="นามสกุล : "+lname,fg="green",font=30)
    textl3.pack()
    textl4=Label(root,text="เลขบัญชีธนาคาร : "+bid,fg="green",font=30)
    textl4.pack()
    textl5=Label(root,text="สาขา : "+area,fg="green",font=30)
    textl5.pack()
        
        
                
    global btn_save,btn_edit,btn_clear
    btn_save = Button(root,text="Save To Excel",command=savetoexcel)
    btn_edit = Button(root,text="Edit",command=edit_info)
    btn_clear = Button(root,text="CLEAR",font=40,command=clearlb)
    btn_save.pack()
    btn_edit.pack()
    btn_clear.pack()

# open excel
def openexcel():
    ex = Tk()

    ex.geometry("500x500") # set the root dimensions
    ex.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
    ex.resizable(0, 0) # makes the root window fixed in size.

    # Frame for TreeView
    frame1 = LabelFrame(ex, text="Excel Data")
    frame1.place(height=250, width=500)



    button2 = Button(ex,text="OPEN EXCEL",width=20,height=5, command=lambda: Load_excel_data())
    button2.place(rely=0.70, relx=0.35)




    ## Treeview Widget
    tv1 = ttk.Treeview(frame1)
    tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

    treescrolly = Scrollbar(frame1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
    treescrollx = Scrollbar(frame1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
    tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
    treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
    treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget
    





    def Load_excel_data():
        file_path = r"C:\\Users\\AAA\\Desktop\\pj ocr\\info.xlsx"
        try:
            excel_filename = r"{}".format(file_path)
            if excel_filename[-4:] == ".csv":
                df = pd.read_csv(excel_filename)
            else:
                df = pd.read_excel(excel_filename)

        except ValueError:
            messagebox.showerror("Information", "The file you have chosen is invalid")
            return None
        except FileNotFoundError:
            messagebox.showerror("Information", f"No such file as {file_path}")
            return None

        clear_data()
        tv1["column"] = list(df.columns)
        tv1["show"] = "headings"
        for column in tv1["columns"]:
            tv1.heading(column, text=column) # let the column heading = column name

        df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in df_rows:
           tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
        return None


    def clear_data():
        tv1.delete(*tv1.get_children())
        return None

#fileselect
def openfile():
    global new_pic,im,my_label,text2
    root.filename = filedialog.askopenfilename(initialdir="",title="select A File",
    filetypes=(("jpg files","*.jpg"),("all files","*.*")))
    filename = root.filename
    #show local file
    text2 = Label(text=filename,fg="blue",font=30)
    text2.pack()
    #show image
    new_pic = ImageTk.PhotoImage(Image.open(filename).resize((225, 300), Image.ANTIALIAS))
    my_label = Label(root,image=new_pic)
    my_label.pack()
    print(type(new_pic))
    #image to scan
    im = cv2.imread(filename)
    im = cv2.resize(im, (500,500))
    flebel.destroy()

def clearlb():
    textl1.destroy()
    textl2.destroy()
    textl3.destroy()
    textl4.destroy()
    textl5.destroy()
    my_label.destroy()
    text2.destroy()
    btn_save.destroy()
    btn_edit.destroy()
    btn_clear.destroy()
#botton text
btn_openexcel = Button(root,text="OPEN EXCEL",font=40,bg="green",command=openexcel).pack()
text = Label(text="Location your photo ID CARD",fg="blue",font=28).pack()
btn_openfile = Button(root,text="Open File",font=40,command=openfile).pack()
btn_scan = Button(root,text="SCAN",bg="blue",font=40,command=scanner).pack()





root.mainloop()