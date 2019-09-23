##Take screenshot of part of the screen
import mss
import mss.tools
from datetime import datetime

with mss.mss() as sct:
    # Get information of monitor 2
    monitor_number = 2
    mon = sct.monitors[monitor_number]
    # The screen part to capture
    monitor = {
        "top": mon["top"] + 115,  # 115px from the top, for normal large screen
        "left": mon["left"] + 320,  # 260px from the left,  for normal large screen
        "width": 1300,              #1300,  for normal large screen
        "height": 700,              #700,  for normal large screen
        "mon": monitor_number,
    }
    dt = datetime.now()
    fname = "pic_{}.png".format(dt.strftime("%H%M_%S"))
        #sct.shot(mon=2, output= "C:\\Users\\tracy.cui\\Desktop\\Screenshot\\" + fname)
    output = "C:\\Users\\tracy.cui\\Desktop\\Screenshot\\" + fname
    # Grab the data
    sct_img = sct.grab(monitor)

    # Save to the picture file
    mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
    
    ##Paste images into powerpoint and save
import pptx
from pptx import Presentation
import pptx.util
import glob
import scipy.misc

OUTPUT_TAG = "F19 BDM Data Review Digital"

# paste in a new powerpoint
prs = pptx.Presentation()
# paste in an existing powerpoint
# prs_exists = pptx.Presentation("Digital_DataReview_2.pptx")

# default slide width
#prs.slide_width = 9144000
# slide height @ 4:3
#prs.slide_height = 6858000
# slide height @ 16:9
prs.slide_height = 5143500

# title slide
slide = prs.slides.add_slide(prs.slide_layouts[6])
# blank slide
#slide = prs.slides.add_slide(prs.slide_layouts[6])

# set title
#title = slide.shapes.title
#title.text = OUTPUT_TAG

pic_left  = int(prs.slide_width * 0.1)
pic_top   = int(prs.slide_height * 0.2)
pic_width = int(prs.slide_width * 0.8)


for g in glob.glob('C:/users/tracy.cui/Desktop/Screenshot/*'):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(50, 50, prs.slide_width,50)  #(left, top, width, height)
    p = tb.text_frame.add_paragraph()
    p.text = "Please Confirm the F19 Data"
    p.font.size = pptx.util.Pt(15)
    p.font.bold = True
    img = scipy.misc.imread(g)
    pic_height = int(pic_width * img.shape[0] / img.shape[1])
    pic = slide.shapes.add_picture(g, pic_left, pic_top, pic_width, pic_height)
    
prs.save('Digital_DataReview_2.pptx')
