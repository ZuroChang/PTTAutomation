# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 16:05:30 2020

@author: ZuroChang
"""

import pandas as pd
# import json


FolderPath='C:/Users/ZuroChang/PythonScript/PTTAutomation/'

class ReadExcel:
    '''
    Description:
        Read the excel file including the ppt content. The output is the format
        required to output the ppt file
    '''
    
    Title=[]
    Pictures=[]
    Content=[]
    SlideLayout=[]
    SlideMapping=[]
    
    def __init__(self,ExcelLocation):
        self.Excel=pd.read_excel(ExcelLocation)\
            .sort_values(by='ID').reset_index(drop=True)
        
        # with open('SlideMapping.json','r') as f:
        #     self.SlideMapping=json.load(f)
    
    def GetSlideMapping(self):
        MappingRef=pd.read_excel(FolderPath+'SlideMapping.xlsx')

        for Template in list(MappingRef['Template'].drop_duplicates()):
            TemplateMap=MappingRef[MappingRef['Template']==Template].reset_index(drop=True)
            
            Mapping=[]
            for i0 in range(TemplateMap.shape[0]):
                Mapping.append(
                    {'Layout':int(TemplateMap['Layout'].iloc[i0])
                    ,'Code':TemplateMap['Code'].iloc[i0]}
                )    
            
            self.SlideMapping.append({'Template':Template,'Mapping':Mapping})
    
    def OutputTitle(self):
        '''
        Description:
            Output self.Title
        '''
        self.Title=list(self.Excel['Title'])
        
    def OutputContent(self):
        '''
        Description:
            Output self.Content
        '''
        def GetBulletLevel(Bullet):
            Bulletlevel={
            'lvl1':[str(i) for i in range(1,100)],
            'lvl2':[chr(i) for i in range(ord('A'),ord('Z'))],
            'lvl3':[chr(i) for i in range(ord('a'),ord('z'))],
            }
            
            for key in Bulletlevel.keys():
                if Bullet in Bulletlevel[key]:
                    return(key)
        
        for i0 in range(self.Excel.shape[0]):
            entry=str(self.Excel['Content'].iloc[i0]).split('\n')
            
            if entry[0]=='nan':
                self.Content.append(None)
            else:
                entryContent=[]
                for i1 in range(len(entry)):
                    Separation=entry[i1].find('. ')
                    Bullet=entry[i1][:Separation]
                    Paragraph=entry[i1][Separation+2:]
                    
                    entryContent.append(
                        {'lvl':GetBulletLevel(Bullet)
                        ,'length':len(Paragraph)
                        ,'Content':Paragraph}
                    )
                
                self.Content.append(entryContent)
        
    def OutputPictures(self):
        '''
        Description:
            Output self.Pictures
        '''
        for i0 in range(self.Excel.shape[0]):
            self.Pictures.append(self.Excel['Pictures'].iloc[i0].split(','))
        
    def OutputSlideLayout(self):
        '''
        Description:
            Output self.Template
        '''
        for i0 in range(self.Excel.shape[0]):
            Template=self.Excel['Template'].iloc[i0]
            for entry in self.SlideMapping:
                if Template==entry['Template']:
                    Mapping=entry['Mapping']
                    break
            
            Code=self.Excel['Slide'].iloc[i0]
            for entry in Mapping:
                if Code==entry['Code']:
                    Layout=entry['Layout']
                    break
                
            self.SlideLayout.append(
                {'Template':Template
                ,'SlideLayout':Layout
                ,'ContentBlock':int(Code[2])}
            )
            
    def Run(self):
        '''
        Description:
            Return the content for outputing ppt file.
        '''
        self.GetSlideMapping()
        self.OutputTitle()
        self.OutputPictures()
        self.OutputContent()
        self.OutputSlideLayout()
        
        return(
            pd.DataFrame(
                {'Title':self.Title
                ,'Pictures':self.Pictures
                ,'Content':self.Content
                ,'SlideLayout':self.SlideLayout}
            )
        )

A1=ReadExcel('C:/Users/ZuroChang/PythonScript/PTTAutomation/ReportGSAM_ZuroChang_20181016.xlsx').Run()
