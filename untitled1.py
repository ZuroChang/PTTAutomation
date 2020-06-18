# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 16:16:42 2020

@author: ZuroChang
"""

import pandas as pd
from pptx import Presentation

PictureFolder='C:/Users/ZuroChang/PythonScript/PTTAutomation/ReportGSAM_ZuroChang_20181016/'
Import=A1
i0=0

Data=Import.iloc[i0]

Template=Data['SlideLayout']['Template']
prs = Presentation(Template+'.pptx')
Slide=prs.slides.add_slide(prs.slide_layouts[Data['SlideLayout']['SlideLayout']]) 



def InsertTitle(Shape,Title):
    Shape.text=Title

def InsertContent(Shape,Content):
    Shape.text=Content

def InsertPicture(Shape,Picture):
    Slide.shapes.add_picture(Picture,Shape.left,Shape.top,Shape.width,Shape.height)

count=0
for shape in Slide.placeholders:
    if 'Title'==shape.name[:len('Title')]:
        InsertTitle(shape,Data['Title'])
    elif 'Picture'==shape.name[:len('Picture')]:
        if count<=len(Data['Pictures']):
            InsertPicture(shape,PictureFolder+Data['Pictures'][count]+'.png')
            count+=1
    # elif 'Content'==shape.name[:len('Content')]:
    #     InsertContent(shape,Data['Content'])

prs.save('test.pptx')
