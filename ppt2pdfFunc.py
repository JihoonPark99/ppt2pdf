import os
import subprocess
from pptx import Presentation
import win32com.client

def ppt2pdf(pptx_file, printtype, blank=True):
    '''
    input
        pptx_file : pdf로 바꿀 ppt의 위치경로
        printtype : int
            1 : 한 페이지에 한 슬라이드
            2 : 한 페이지에 두 슬라이드
            3 : 한 페이지에 세 슬라이드 (연습장도 같이 있고, 
                이땐, 중간중간 빈칸(blank) 안만들어줘도 될것 같음.)
            4 : 한 페이지에 여섯 페이지
            5 : 한 페이지에 한 슬라이드 (작게)
            6 : 개요
        blank : True(default) - 중간중간 빈칸을 만듬
                False - 빈칸을 만들지 않음.
    '''
    # PowerPoint 프로세스를 시작합니다.
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True

    # pptx_file열기
    presentation = powerpoint.Presentations.Open(pptx_file)
    slides = presentation.Slides
    
    if blank == True : 
        ## 중간중간 빈칸 만들어주기
        posi_li = [x*2 for x in range(1, len(slides)+1)]
        for position in posi_li:

            slide_layout = slides[position - 1].CustomLayout
            new_slide = slides.AddSlide(position, slide_layout)
    else : 
        pass

    presentation.PrintOptions.OutputType = printtype

    ## 인쇄할 사본의 수
    presentation.PrintOptions.NumberOfCopies = 1

    presentation.PrintOut()
#     presentation.Close()
#     powerpoint.Quit()