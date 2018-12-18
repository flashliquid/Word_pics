'''this will insert all the images in a directory into a table in a word document, the "total column" value can be adjusted
to create multiple columns if needed.  update with appropriate directory info "my _dir" '''
import win32com.client as win32
import os
pathx=raw_input("Please input the locationo of the folder containing the images:")
pathx =  pathx.replace( "\\", "/")#avoid escape character issues

my_dir = pathx[1:len(pathx)-1]
file_type =raw_input("Please enter the file extension, for example for .jpg just enter jpg:")
file_type = file_type.upper()
#creating a word application object
total_column= raw_input("How many columns of pictures in the document ?")
wordApp = win32.gencache.EnsureDispatch('Word.Application') #create a word application object
wordApp.Visible = True # hide the word application
doc = wordApp.Documents.Add() # create a new application

#Formating the document
doc.PageSetup.RightMargin = 10
doc.PageSetup.LeftMargin = 10
doc.PageSetup.Orientation = win32.constants.wdOrientLandscape
# a4 paper size: 595x842
doc.PageSetup.PageWidth = 595
doc.PageSetup.PageHeight = 842


# Inserting Tables
#my_dir="C:/Users/David/Documents/EGi/EGi Plots/ARD_plots/ARD_boxplots/Current_sites"
filenames = os.listdir(my_dir)
piccount=0
file_count = 0
for i in filenames:
    if i[len(i)-3: len(i)].upper() == file_type: # check whether the current object is a JPG file
        piccount = piccount + 1
        #file_count= file_count + 1
print piccount, " images will be inserted"
#print filenames
total_column = int(total_column)
total_row = int(piccount/total_column)+2
rng = doc.Range(0,0)
rng.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
table = doc.Tables.Add(rng,total_row, total_column)
table.Borders.Enable = False
if total_column > 1:
    table.Columns.DistributeWidth()

#Collecting images in the same directory and inserting them into the document
#frame_max_width= 167 # the maximum width of a picture
#frame_max_height= 125 # the maximum height of a picture


piccount = -1

for index, filename in enumerate(filenames): # loop through all the files and folders for adding pictures
    #if os.path.isfile(os.path.join(os.path.abspath("."), filename)):
    if os.path.isfile(os.path.join(os.path.abspath(my_dir), filename)): # check whether the current object is a file or not
        if filename[len(filename)-3: len(filename)].upper() == file_type: # check whether the current object is a JPG file
            piccount = piccount + 1
            #print filename, len(filename), filename[len(filename)-3: len(filename)].upper()


            cell_column = (piccount % total_column + 1) #calculating the position of each image to be put into the correct table cell
            cell_row = (piccount/total_column + 1)
            #print 'cell_column=%s,cell_row=%s' % (cell_column,cell_row)

            #we are formatting the style of each cell
            cell_range= table.Cell(cell_row, cell_column).Range
            cell_range.ParagraphFormat.LineSpacingRule = win32.constants.wdLineSpaceSingle
            cell_range.ParagraphFormat.SpaceBefore = 0
            cell_range.ParagraphFormat.SpaceAfter = 3

            #this is where we are going to insert the images
            current_pic=cell_range.InlineShapes.AddPicture(os.path.join(os.path.abspath(my_dir), filename))
            #width, height = (frame_max_height*frame_max_width/frame_max_height, frame_max_height)

            #changing the size of each image to fit the table cell
            #current_pic.Height= height
            #current_pic.Width= width
            #putting a name underneath each image which can be handy
            current_pic.Range.InsertCaption(Label="Figure", Title=": "+filename[:len(filename)-4]+"\n\n\n")
            #table.Cell(cell_row, cell_column).Range.InsertAfter("\n"+"Appendix II, Figure "+ str(piccount-1)+": "+filename[:len(filename)-4]+"\n"+"\n"+"\n")
        else: continue
