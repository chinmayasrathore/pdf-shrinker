# -*- coding: utf-8 -*-

################################################################################
# Version 1.00 Dr. Chinmaya S Rathore Oct' 2025 
# PDF Shrinker is a simple and easy to use GUI that runs a Ghostscript Command
# to shrink or compress PDF files. It requires that Ghostscript is installed.
# Ghostscript is available from https://www.ghostscript.com/releases/gsdnld.html
################################################################################


from PySide6.QtCore import (QCoreApplication, QMetaObject, QThread,Signal, QTimer, QSize, Qt)
from PySide6.QtGui import (QFont, QIcon,QMovie)
from PySide6.QtWidgets import (QApplication, QComboBox, QFrame, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QSizePolicy,QMessageBox,QFileDialog,
    QToolButton, QVBoxLayout, QWidget, QSpinBox)
from datetime import datetime
import os
import subprocess
import shutil


#Get app folder path. Used for creating path to icon files in UI Form creation 
basedir = os.path.dirname(__file__)


class Worker(QThread):
    finished = Signal(int,str,str)
    error = Signal(str)

    def __init__(self,command):
        super().__init__()
        self.command = command
        
    def run(self):
        try:
            #Run the GhostScript command         
            result=subprocess.run(self.command, shell=True, capture_output=True, text=True, check=True)
            self.finished.emit(result.returncode, result.stdout,result.stderr)
        
        except subprocess.CalledProcessError as err:
            # Handle errors from the subprocess itself (e.g., command not found, non-zero exit code)
            self.error.emit(f"{err.stderr or str(err)}")
        
        except Exception as exc:  
            self.error.emit(str(exc)) 
            
class Ui_Form(object):
    def setupUi(self, Form):
        # Test if Ghostscript is installed and if yes set self.gs_version to the installed version
        # Call gs_installed() which sets the instance variable self.gs_version if installed or exits.
        self.gs_version=None
        self.gs_installed()

        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.resize(781, 536)
        icon = QIcon(os.path.join(basedir,"icons","complogo.png")) # Application Logo Icon - Used complogo48.svg for Linux builds
        Form.setWindowIcon(icon)
        self.verticalLayout = QVBoxLayout(Form)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.label_inputPDF = QLabel(Form)
        self.label_inputPDF.setObjectName(u"labelInputFile")
        self.label_inputPDF.setMaximumSize(QSize(125, 16777215))
        font = QFont()
        font.setPointSize(12)
        self.label_inputPDF.setFont(font)

        self.horizontalLayout.addWidget(self.label_inputPDF)

        self.lineEdit_InputPDF = QLineEdit(Form)
        self.lineEdit_InputPDF.setObjectName(u"lineEdit_InputPDF")
        self.lineEdit_InputPDF.setFont(font)
        #ensuring that user selects a PDF file and does not type in a name
        self.lineEdit_InputPDF.setReadOnly(True)

        self.horizontalLayout.addWidget(self.lineEdit_InputPDF)

        self.pushButton_browse = QPushButton(Form)
        self.pushButton_browse.setObjectName(u"pushButton_browse")
        font1 = QFont()
        font1.setPointSize(10)
        self.pushButton_browse.setFont(font1)
        #Hover Color on Push Button Browse and similar on other push buttons
        self.pushButton_browse.setStyleSheet("""
                                             QPushButton {background-color:wheat; 
                                             color:black;}
                                             QPushButton::hover{
                                             background-color: #0077e6;
                                             color:white}""")
        
        self.input_pdf = ""  # will contain the path to input pdf file
        self.horizontalLayout.addWidget(self.pushButton_browse)
        
    # Browse Button Signal Call 
        self.pushButton_browse.clicked.connect(self.browse_file) # Button Clicked

        self.verticalLayout.addLayout(self.horizontalLayout)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.Label_OutputPDF = QLabel(Form)
        self.Label_OutputPDF.setObjectName(u"label_OutputFile")
        self.Label_OutputPDF.setMaximumSize(QSize(115, 16777215))
        self.Label_OutputPDF.setFont(font)

        self.horizontalLayout_2.addWidget(self.Label_OutputPDF)

        self.lineEdit_outputPDF = QLineEdit(Form)
        self.lineEdit_outputPDF.setObjectName(u"lineEdit_outputPDF")
        self.lineEdit_outputPDF.setFont(font)
        self.lineEdit_outputPDF.setReadOnly(True)

        self.horizontalLayout_2.addWidget(self.lineEdit_outputPDF)

        self.pushButton_saveAs = QPushButton(Form)
        self.pushButton_saveAs.setObjectName(u"pushButton_saveAs")
        font1 = QFont()
        font1.setPointSize(10)
        self.pushButton_saveAs.setFont(font1)
        self.pushButton_saveAs.setStyleSheet("""
                                             QPushButton {background-color:wheat; 
                                             color:black;}
                                             QPushButton::hover{
                                             background-color: #0077e6;
                                             color:white}""")
        
        self.output_pdf = ""  # will hold the save as or output file name and path

        self.horizontalLayout_2.addWidget(self.pushButton_saveAs)
        
        # Save As Button Signal Call 
        self.pushButton_saveAs.clicked.connect(self.save_file) # saveAs Button Clicked


        self.verticalLayout.addLayout(self.horizontalLayout_2)

        # Compression Option Label 
        self.horizontalLayout_3 = QHBoxLayout()
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.Label_Compression = QLabel(Form)
        self.Label_Compression.setObjectName(u"LabelCompression")
        self.Label_Compression.setMaximumSize(QSize(115, 16777215))
        self.Label_Compression.setFont(font)

        self.horizontalLayout_3.addWidget(self.Label_Compression)

        # Compression Option Combobox
        self.comboBox_compression = QComboBox(Form)
        self.comboBox_compression.addItem("")
        self.comboBox_compression.addItem("")
        self.comboBox_compression.addItem("")
        self.comboBox_compression.addItem("")
        self.comboBox_compression.addItem("")
        self.comboBox_compression.setObjectName(u"comboBox_compression")
        self.comboBox_compression.setFont(font)

        self.horizontalLayout_3.addWidget(self.comboBox_compression)

        # Compression option Information Button 

        self.toolButton_Compression = QToolButton(Form)
        self.toolButton_Compression.setObjectName(u"toolButton_Compression")
        self.toolButton_Compression.setToolTipDuration(-1)
        icon_info = QIcon(os.path.join(basedir,"icons","inforsvg52.svg"))
        self.toolButton_Compression.setIcon(icon_info)
        self.toolButton_Compression.setIconSize(QSize(25, 25))
        self.toolButton_Compression.setPopupMode(QToolButton.DelayedPopup)
        self.toolButton_Compression.setAutoRaise(True)

        self.horizontalLayout_3.addWidget(self.toolButton_Compression)

        self.verticalLayout.addLayout(self.horizontalLayout_3)

        # Compatibility Option Label 
        self.horizontalLayout_4 = QHBoxLayout()
        self.horizontalLayout_4.setObjectName(u"horizontalLayout_4")
        self.Label_Compatibility = QLabel(Form)
        self.Label_Compatibility.setObjectName(u"LabelCompatibility")
        self.Label_Compatibility.setMaximumSize(QSize(115, 16777215))
        self.Label_Compatibility.setFont(font)

        self.horizontalLayout_4.addWidget(self.Label_Compatibility)
        
        # Compatibility Option Combo Box 
        self.comboBox_compatbility = QComboBox(Form)
        self.comboBox_compatbility.addItem("")
        self.comboBox_compatbility.addItem("")
        self.comboBox_compatbility.addItem("")
        self.comboBox_compatbility.addItem("")
        self.comboBox_compatbility.addItem("")
        self.comboBox_compatbility.addItem("")
        self.comboBox_compatbility.setObjectName(u"comboBox_compatbility")
        self.comboBox_compatbility.setFont(font)

        self.horizontalLayout_4.addWidget(self.comboBox_compatbility)

        # Compatibility Option Information Button 
        self.toolButton_compatibility = QToolButton(Form)
        self.toolButton_compatibility.setObjectName(u"toolButton_compatibility")
        self.toolButton_compatibility.setIcon(icon_info)
        self.toolButton_compatibility.setIconSize(QSize(25, 25))
        self.toolButton_compatibility.setPopupMode(QToolButton.DelayedPopup)
        self.toolButton_compatibility.setAutoRaise(True)

        self.horizontalLayout_4.addWidget(self.toolButton_compatibility)


        self.verticalLayout.addLayout(self.horizontalLayout_4)

        # Advanced Toggle Button 
        self.advanced_btn = QToolButton()
        self.advanced_btn.setText("Advanced")
        self.advanced_btn.setFont(font)
        self.advanced_btn.setCheckable(True)

        self.verticalLayout.addWidget(self.advanced_btn)

        self.adv_widget = QWidget()
        self.adv_widget.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        adv_layout = QVBoxLayout()    # Vertical layout
        adv_layout.setSpacing(25)
        adv_layout.setContentsMargins(0,0,0,0)
        advh_layout1 = QHBoxLayout()   # first horizontal layout
        advh_layout1.setSpacing(25)
        advh_layout1.setContentsMargins(0,0,0,0)

        # Advanced Section Image Resolution Label 
        self.label_cir= QLabel()
        self.label_cir.setText("Image Resolution:")
        self.label_cir.setFont(font)
        advh_layout1.addWidget(self.label_cir)    # add to horizontal layout1 

        # Advanced Section Spinbox for image resolution (user input) 
        self.spinbox_Resolution = QSpinBox()
        self.spinbox_Resolution.setFont(font)
        self.spinbox_Resolution.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.spinbox_Resolution.setRange(45,250)
        self.spinbox_Resolution.setSingleStep(5)
        self.spinbox_Resolution.setSpecialValueText("Default")
        self.spinbox_Resolution.setValue(45) # An initial value is set when no choice is made by user (not used)
        advh_layout1.addWidget(self.spinbox_Resolution)    # add to horizontal layout1 


        #Information Button for Image Resolution
        self.toolButton_cir = QToolButton(Form)
        self.toolButton_cir.setObjectName(u"toolButton_Compression")
        self.toolButton_cir.setToolTipDuration(-1)
        self.toolButton_cir.setIcon(icon_info)
        self.toolButton_cir.setIconSize(QSize(25, 25))
        self.toolButton_cir.setPopupMode(QToolButton.DelayedPopup)
        self.toolButton_cir.setAutoRaise(True)

        advh_layout1.addWidget(self.toolButton_cir)  # add to horizontal layout1 

        # Advanced Section Image Resolution information button Signal call 
        self.toolButton_cir.clicked.connect(self.show_cir_help)   #Image Resolution information button Clicked

        # Advanced Section Downsample Option Label 

        adv_layout.addLayout(advh_layout1) # add horizontal layout 1 to vertical layout adv_layout
        
        advh_layout2 = QHBoxLayout()  # second horizontal layout
        self.label_dt= QLabel()
        self.label_dt.setText("Downsample Threshold:")
        self.label_dt.setFont(font)
        advh_layout2.addWidget(self.label_dt)     # add to horizontal layout2 
        
        # Advanced Section Downsample Option Combo Box (user input)

        advh_layout2.setSpacing(25)
        self.combo_qfactor=QComboBox()
        self.combo_qfactor.setFont(font)
        self.combo_qfactor.setMaximumSize(QSize(115, 16777215))
        self.combo_qfactor.addItems(['Default','1','2','3','4'])
        advh_layout2.addWidget(self.combo_qfactor)   # add to horizontal layout2 


        #Advanced Section Information  Button for Downsample Threshold
        self.toolButton_dt = QToolButton(Form)
        self.toolButton_dt.setObjectName(u"toolButton_Compression")
        self.toolButton_dt.setToolTipDuration(-1)
        self.toolButton_dt.setIcon(icon_info)
        self.toolButton_dt.setIconSize(QSize(25, 25))
        self.toolButton_dt.setPopupMode(QToolButton.DelayedPopup)
        self.toolButton_dt.setAutoRaise(True)

        advh_layout2.addWidget(self.toolButton_dt)  # add toolButton to horizontal layout2

        # Downsample Factor Combo Information Button Signal Call 
        self.toolButton_dt.clicked.connect(self.show_dt_help)   #Button Clicked

        adv_layout.addLayout(advh_layout2)  # add horizontal layout 2 to vertical layout adv_layout 


        # Advanced Section - Grayscale Option Label 
        advh_layout3 = QHBoxLayout()  # Third horizontal layout
        self.label_grayscale= QLabel()
        self.label_grayscale.setText("Convert PDF to Grayscale:")
        self.label_grayscale.setFont(font)
        advh_layout3.addWidget(self.label_grayscale) # Add to horizontal layout 3
        
        # Advanced Section - Grayscale Combo Box ( user input) 
        advh_layout3.setSpacing(25)
        self.combo_grayscale=QComboBox()
        self.combo_grayscale.setFont(font)
        self.combo_grayscale.setMaximumSize(QSize(115, 16777215))
        self.combo_grayscale.addItems(['No','Yes'])
        advh_layout3.addWidget(self.combo_grayscale)  # Add to horizontal layout 3


        #Information Button for GrayScale Option in Advanced Section 
        self.toolButton_grayscale = QToolButton(Form)
        self.toolButton_grayscale.setObjectName(u"toolButton_Grayscale")
        self.toolButton_grayscale.setToolTipDuration(-1)
        self.toolButton_grayscale.setIcon(icon_info)
        self.toolButton_grayscale.setIconSize(QSize(25, 25))
        self.toolButton_grayscale.setPopupMode(QToolButton.DelayedPopup)
        self.toolButton_grayscale.setAutoRaise(True)

        advh_layout3.addWidget(self.toolButton_grayscale) # Add to horizontal layout 3

        # Downsample Factor Combo Information Button Signal Call 
        self.toolButton_grayscale.clicked.connect(self.show_grayscale_help)   #Grayscale info button Clicked

        adv_layout.addLayout(advh_layout3)  # add horizontal layout3 to vertical layout adv_layout 
        #Advanced Section UI Ends

        # Add the vertical layout to the widget adv_widget

        self.adv_widget.setLayout(adv_layout)  
        self.adv_widget.setVisible(False)
        self.verticalLayout.addWidget(self.adv_widget)

        # Advanced Button Toggle 
        self.advanced_btn.toggled.connect(self.adv_widget.setVisible) #Hide/Unhide Advanced Section when Advanced button clicked

        # Finishing Advanced Section 
        self.line = QFrame(Form)
        self.line.setObjectName(u"line")
        self.line.setFrameShape(QFrame.Shape.HLine)
        self.line.setFrameShadow(QFrame.Shadow.Sunken)

        self.verticalLayout.addWidget(self.line)

        # Action Buttons Row that has Shrink PDF, Clear , Help and Quit Buttons 
        self.horizontalLayout_5 = QHBoxLayout()
        self.horizontalLayout_5.setObjectName(u"horizontalLayout_5")
        self.pushButton_compress = QPushButton(Form)
        self.pushButton_compress.setObjectName(u"pushButton_compress")
        self.pushButton_compress.setFont(font)
        self.pushButton_compress.setStyleSheet("""
                                             QPushButton {background-color:#ffe6e6; 
                                             color:black;}
                                             QPushButton::hover{
                                             background-color: #0077e6;
                                             color:white}""")

        self.horizontalLayout_5.addWidget(self.pushButton_compress) # Add Shrink PDF Button to Horizontal layout5

        self.pushButton_clear = QPushButton(Form)
        self.pushButton_clear.setObjectName(u"pushButton_clear")
        self.pushButton_clear.setFont(font)
        self.pushButton_clear.setStyleSheet("""
                                             QPushButton {background-color:#ffe6e6; 
                                             color:black;}
                                             QPushButton::hover{
                                             background-color: #0077e6;
                                             color:white}""")

        self.horizontalLayout_5.addWidget(self.pushButton_clear)  # Add Clear button to Horizontal layout5

        # Clear Button Signal Call
        self.pushButton_clear.clicked.connect(self.clear_fields)  # Clear Button_clicked


        #PushButton Help 
        self.pushButton_help = QPushButton(Form)
        self.pushButton_help.setObjectName(u"pushButton_help")
        self.pushButton_help.setFont(font)
        self.pushButton_help.setStyleSheet("""
                                             QPushButton {background-color:#ffe6e6; 
                                             color:black;}
                                             QPushButton::hover{
                                             background-color: #0077e6;
                                             color:white}""")
        self.horizontalLayout_5.addWidget(self.pushButton_help) # Add Help Button to Horizontal layout5

       # Help PushButton Signal Call
        self.pushButton_help.clicked.connect(self.show_help) # Help Button Clicked

        #PushButton Quit
        self.pushButton_quit = QPushButton(Form)
        self.pushButton_quit.setObjectName(u"pushButton_log")
        self.pushButton_quit.setFont(font)
        self.pushButton_quit.setStyleSheet("""
                                             QPushButton {background-color:#ffe6e6; 
                                             color:black;}
                                             QPushButton::hover{
                                             background-color: #0077e6;
                                             color:white}""")
        self.horizontalLayout_5.addWidget(self.pushButton_quit) # Add Quit Button to Horizontal layout5

        
        self.verticalLayout.addLayout(self.horizontalLayout_5)   # Add Horizontal layout5 to vertical layout

        # Activity or Status Messages Label  
        self.horizontalLayout_6 = QHBoxLayout()
        self.horizontalLayout_6.setObjectName(u"horizontalLayout_6")

        self.label_5 = QLabel(Form)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setFont(font)
        self.label_5.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.label_5.setMinimumHeight(25)
        self.label_5.setMaximumHeight(25)
        self.label_5.setTextFormat(Qt.RichText)         # ensure HTML is interpreted
        self.label_5.setTextInteractionFlags(Qt.TextBrowserInteraction)  # allow clicking & selecting
        self.label_5.setOpenExternalLinks(True)    
        self.label_5.setText("🤓Welcome to PDF Shrinker. Let's shrink some PDFs...")
        self.label_5.setStyleSheet("background-color:#cce6ff;")

        self.horizontalLayout_6.addWidget(self.label_5)  # Add Label to Horizontal Layout6

       # Add Loader GIF Label ( This GIF is used to show working status to user)
       # Shown only when Ghostscript command runs 

        self.loader_label = QLabel(Form)
        self.loader_label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.loader_label.setMinimumHeight(50)
        self.loader_label.setMaximumHeight(50)
        self.horizontalLayout_6.addWidget(self.loader_label)
        self.movie=QMovie(os.path.join(basedir,"icons","loader.gif"))
        self.movie.setScaledSize(QSize(48,48))
        self.loader_label.setMovie(self.movie)
        self.movie.start()
        self.loader_label.hide()  # Hiding Loader GIF Initially. self.loader_label.show() to show
        #End Loader GIF addition and display 

        self.horizontalLayout_6.addWidget(self.loader_label)   # Add Loader label to Horizontal Layout6
        self.verticalLayout.addLayout(self.horizontalLayout_6) # Add Horizontal Layout6 to Vertical Layout
        # UI Form elements done! 

        self.retranslateUi(Form)
        QMetaObject.connectSlotsByName(Form)

    # setup Ui
    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"PDF Shrinker - Version 1.00 (4-10-2025)  - Dr. Chinmaya S Rathore", None))
        self.label_inputPDF.setText(QCoreApplication.translate("Form", u"Input PDF", None))
        # self.lineEdit_InputPDF.setToolTip("")
        self.lineEdit_InputPDF.setPlaceholderText(QCoreApplication.translate("Form", u"<Please Select a PDF file to compress>", None))
        self.pushButton_browse.setText(QCoreApplication.translate("Form", u"Browse", None))
        self.Label_OutputPDF.setText(QCoreApplication.translate("Form", u"Output PDF", None))
        self.pushButton_saveAs.setText(QCoreApplication.translate("Form", u"Save As", None))
        self.lineEdit_outputPDF.setPlaceholderText(QCoreApplication.translate("Form", u"<output pdf name ex. output.pdf>", None))
        self.Label_Compression.setText(QCoreApplication.translate("Form", u"Compression", None))
        self.comboBox_compression.setItemText(0, QCoreApplication.translate("Form", u"default", None))
        self.comboBox_compression.setItemText(1, QCoreApplication.translate("Form", u"screen", None))
        self.comboBox_compression.setItemText(2, QCoreApplication.translate("Form", u"ebook", None))
        self.comboBox_compression.setItemText(3, QCoreApplication.translate("Form", u"printer", None))
        self.comboBox_compression.setItemText(4, QCoreApplication.translate("Form", u"prepress", None))
        
        # Compression Combo Information Button Signal Call 
        self.toolButton_Compression.clicked.connect(self.show_dialog_levels)   #Button Clicked

        self.toolButton_Compression.setText(QCoreApplication.translate("Form", u"...", None))
        self.Label_Compatibility.setText(QCoreApplication.translate("Form", u"Compatibility", None))
        self.comboBox_compatbility.setItemText(0, QCoreApplication.translate("Form", u"1.3", None))
        self.comboBox_compatbility.setItemText(1, QCoreApplication.translate("Form", u"1.4", None))
        self.comboBox_compatbility.setItemText(2, QCoreApplication.translate("Form", u"1.5", None))
        self.comboBox_compatbility.setItemText(3, QCoreApplication.translate("Form", u"1.6", None))
        self.comboBox_compatbility.setItemText(4, QCoreApplication.translate("Form", u"1.7", None))
        self.comboBox_compatbility.setItemText(5, QCoreApplication.translate("Form", u"2.0", None))
        self.comboBox_compatbility.setCurrentIndex(1)
        self.toolButton_compatibility.setText(QCoreApplication.translate("Form", u"...", None))

        self.pushButton_compress.setText(QCoreApplication.translate("Form", u"Shrink PDF", None))
        self.pushButton_clear.setText(QCoreApplication.translate("Form", u"Clear", None))
        self.pushButton_help.setText(QCoreApplication.translate("Form", u"Help", None))
        self.pushButton_quit.setText(QCoreApplication.translate("Form", u"Quit", None))


    # Compatibility Information Button Signal Call 
        self.toolButton_compatibility.clicked.connect(self.show_dialog_compatibility)   #Compatibility Information Button Clicked   

    # Quit Button Clicked -> Exit Application 
        self.pushButton_quit.clicked.connect(self.exit_app) 
    
    # Shrink PDF Button Signal Call 
        self.pushButton_compress.clicked.connect(self.on_run)  # Shrink PDF button clicked

    # Information button slot for PDF compression setting options
    def show_dialog_levels(self):
        QMessageBox.information(MainForm,"Compression Levels Help","Default - Moderate file size - Moderate Quality\nScreen - Smallest file size - Lower Quality\neBook - Moderate file size - Better quality\nPrint- Larger file size - High quality\nPreepress-Large file size- Best Quality\n\nQuality largely affects images in the PDF and you can try different levels to see what works best.")
     
    # Information button slot for Acrobat compatibility setting options
    def show_dialog_compatibility(self):
        QMessageBox.information(MainForm,"Acrobat Compatibility Help","This parameter defines " \
                                         "compatibility with different Acrobat Versions:\n\n 1.4 "
                                         "- Compatible with Acrobat 5 and later (Safe Default)\n " \
                                         "1.5 - Compatible with Acrobat 6 and later\n " \
                                         "1.6 - Compatible with Acrobat 7 and later\n 1.7 - " \
                                         "Compatible with Acrobat 8 and later\n 2.0 - Compatible" \
                                         " with Acrobat 10 and later\n 1.3 - Compatible with " \
                                         "Acrobat 3 & later (Flattens PDF-Removes Transparencies)" \
                                         "\n\n 1.5, 1.6, 1.7 or 2.0 may offer better compression.")
       
    # Get Input File name browse button dialog slot
    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(caption="Select PDF file",dir="",filter="PDF Files (*.pdf)")
        if path:
            self.lineEdit_InputPDF.setText(path)

    #Get Output file name browse button dialog slot
    def save_file(self):
         path, _ = QFileDialog.getSaveFileName(caption="Save Output PDF file as",dir="",filter="PDF Files (*.pdf)")
         if path:
            self.lineEdit_outputPDF.setText(path)

    # Information button slot for color image resolution setting options
    def show_cir_help(self):
        QMessageBox.information(MainForm,"Image Resolution Help","This applies to images in the PDF. " \
        "It permits you to choose a finer level of resolution when images are resampled during compression " \
        "overriding the preset resolution as chosen in the compression setting.\n\n " \
        "A setting of 95-120 when compression is set to ebook gives smaller yet screen readable PDF files. " \
        "You can experiment with lower or higher resolutions. Verify the quality of images in the output PDF. ")

    # Information button slot for down sample factor setting options
    def show_dt_help(self):
        QMessageBox.information(MainForm,"Downsample Threshold Help","This setting permits decides what resolution of images in the PDF will be downsampled.This setting works in conjugation with Color Image Resolution setting.\n\nChoosing 1 achieves smallest file size.")


    # Information button slot for grayscale setting options
    def show_grayscale_help(self):
        QMessageBox.information(MainForm,"Convert to Grayscale Help","Choosing 'Yes' for this option " \
        "converts the color PDF to a Grayscale PDF significantly reducing its size. " \
        "All color images are converted to Grayscale images. \n\n" \
        "Useful for text-centric PDFs like legal, academic, forms, invoices, letters, contracts, " \
        "reports not needing color illustrations and documents for OCR etc. Also lowers printing cost.")

    #===================Start Compress PushButton slot on_run

    # Slot for the Compress Push Button 
    def on_run(self):
        
      # Get input and output filenames from the lineEdit fields
        input_pdf = self.lineEdit_InputPDF.text().strip()
        output_pdf = self.lineEdit_outputPDF.text().strip() 
     
      # Check if input file name has been assigned
        if not input_pdf:
           QMessageBox.warning(MainForm,"No Input File Chosen","Please choose an input PDF file to compress.")
           self.lineEdit_InputPDF.clear()
           return
            
        # Check if output file name has been assigned
        if not output_pdf:
           QMessageBox.warning(MainForm,"No Output File Name Provided","Please choose or enter a name for the output file!")
           self.lineEdit_outputPDF.clear()
           return
        
        # Compare input and output file names to ensure that input PDF cannot be overwritten 
        if input_pdf == output_pdf:
            QMessageBox.critical(MainForm, "Cannot Overwrite Input PDF File", "Input file and output file cannot be the same!\nPlease try again with a different output file name.")
            self.lineEdit_outputPDF.clear()
            return
           
        # 1. Assigning chosen user inputs for compression and compatibility to variables which will be used to concatenate the ghostscript terminal command
        
        quality = self.comboBox_compression.currentText().strip()
        compatibility = self.comboBox_compatbility.currentText().strip()

        # 2. Assign user inputs from the advanced section if chosen to be used to concatenate the 
          # ghostscript terminal command parse advance returns a tuple and individual 
          # values in the tuple are assigned to variables in sequence 
          # example a,b,c,d = [10,20,30,40] > a=10, b=20, c=30, d=40
        
        image_res,image_dst,image_ds_true,grayscale_params = self.parse_advance()
        
        # 3. Concatenate GhostScript Command adding options chosen by user for compression , resolution etc.
        command = f'{self.gs_version} -sDEVICE=pdfwrite -dNOPAUSE -dQUIET -dBATCH -dSAFER -dCompatibilityLevel={compatibility} -dPDFSETTINGS=/{quality} ' + f'{image_ds_true}' + f'{image_res} ' + f'{image_dst} ' + f'{grayscale_params} ' +f'-sOutputFile="{output_pdf}" ' + f'"{input_pdf}"'
        
        # We now run the Ghostscript terminal command in a different thread
        self.run_compress(command)
    #=======================End Compress PushButton slot on_run

    # Parsing and setting variables for Advanced inputs if opted by user
    def parse_advance(self):
        var_cir=self.spinbox_Resolution.value()
        var_dst=self.combo_qfactor.currentText().strip()
        var_gray=self.combo_grayscale.currentText().strip()
        if var_cir == 45 and var_dst == "Default":
            image_res=""
            image_dst=""
            image_ds_true=""
        elif var_cir != 45 and var_dst == "Default":
            image_res=" -dColorImageResolution="+ str(var_cir)
            image_dst=""
            image_ds_true=" -dDownsampleColorImages=true "
        elif var_cir == 45 and var_dst != "Default":
            image_res=""
            image_dst=" -dColorImageDownsampleThreshold="+ str(var_dst)
            image_ds_true=" -dDownsampleColorImages=true "
        elif var_cir != 45 and var_dst != "Default":
            image_res=" -dColorImageResolution="+ str(var_cir)
            image_dst=" -dColorImageDownsampleThreshold="+ str(var_dst)
            image_ds_true=" -dDownsampleColorImages=true "

        # Assess Grayscale Options with downsample threshold  and resolution choices
        # If var_cir = 45 => user has not chosen any image resolution (default option) 

        if   var_gray == "No":                  
           grayscale_params=""
        
        elif var_gray == "Yes" and var_dst == "Default" and var_cir == 45:     # for resolution, cir is = 45 when user does not make any choice i.e. cir is default
           grayscale_params=f"-sColorConversionStrategy=Gray -dProcessColorModel=/DeviceGray "            
            
        elif var_gray == "Yes" and var_dst != "Default" and var_cir != 45:
           grayscale_params=f"-sColorConversionStrategy=Gray -dProcessColorModel=/DeviceGray -dGrayImageDownsampleThreshold={str(var_dst)} -dGrayImageResolution={var_cir} -dMonoImageResolution={var_cir} -dMonoImageDownsampleThreshold={str(var_dst)} " 

        elif var_gray == "Yes" and var_dst == "Default" and var_cir != 45:
           grayscale_params=f"-sColorConversionStrategy=Gray -dProcessColorModel=/DeviceGray -dGrayImageResolution={var_cir} -dMonoImageResolution={var_cir} " 
    
        elif var_gray == "Yes" and var_dst != "Default" and var_cir == 45:
           grayscale_params=f"-sColorConversionStrategy=Gray -dProcessColorModel=/DeviceGray -dGrayImageDownsampleThreshold={str(var_dst)} -dMonoImageDownsampleThreshold={str(var_dst)} " 
            
        return image_res,image_dst,image_ds_true,grayscale_params 
    
    # Test if Ghostscript is installed. If yes return the version else return 0 
    def gs_installed(self):    
        # Find out if (a) Ghostscript is installed and (b) the Version if installed
        
        # set the instance gs_version variable if Ghostscript is installed. This will be used in concatenating command
        for exe in ("gswin64c.exe", "gswin32c.exe","gs"):
            mpath = shutil.which(exe)
            if mpath:
              full_name = os.path.basename(mpath)
              self.gs_version = os.path.splitext(full_name)[0]
              break  # Exit For loop when found - i.e. stop further searching in exe
              
        # Message to user to install Ghostscript before using PDF Shrinker       
        if not self.gs_version:
            QMessageBox.critical(MainForm,"Ghostscript Not Installed","Please "+'<a href="https://ghostscript.com/releases/gsdnld.html"> '\
            'download & install Ghostscript</a>' +". Choose a 64-bit or 32-bit release as appropriate for \
            your machine."+ "\nThis is a prerequisite for using this software.\nGhostScript is free and \
            open source!\n\nPlease try again after installing Ghostscript!")
            sys.exit()
        return  
    
# Run the Ghostscript command in the terminal as a separate thread  (GUI is active to show Loader GIF)
    def run_compress(self,command):
        # Deactivate Push button Shrink PDF - user cannot press this button while thread is running 
        self.pushButton_compress.setEnabled(False)  # Prevent starting another thread
        self.pushButton_clear.setEnabled(False) # Prevent reset of input and output file names
        self.pushButton_quit.setEnabled(False) #Prevent closing app while thread is running 
        self.label_5.setStyleSheet("background-color:#ffe6cc;")
        self.label_5.setText("🏋️Working...Large files and Grayscale conversion may take some time! Please Wait...")

        #Start Spinner GIF 
        self.loader_label.show()
        self.movie.start()

        # Create worker object and start the worker thread to run the command 
        self.worker = Worker(command)
        self.worker.finished.connect(self.handle_result)
        self.worker.error.connect(self.handle_error)
        self.worker.start()


# Change status message after a successful run 
    def clear_status(self): 
        self.label_5.setStyleSheet("background-color:#cce6ff;")
        self.label_5.setText("👉 Ready again! Choose Different Compression Options, Clear & Start Over, See Help or Quit. Hmm...🤔")

# Reset all user input fields and restore default settings in combo boxes and spinbox
    def clear_fields(self):
        #reset file name input and output boxes
        self.lineEdit_InputPDF.clear()
        self.lineEdit_outputPDF.clear()
        #reset combo boxes
        self.comboBox_compression.setCurrentIndex(0)
        self.comboBox_compatbility.setCurrentIndex(1)
        
        #reset Advanced resolution, downsample threshold and grayscale options to defaults
        self.spinbox_Resolution.setValue(0)
        self.combo_qfactor.setCurrentIndex(0)
        self.combo_grayscale.setCurrentIndex(0)
       
        # Clear Advanced Button and Collapse Advanced Options as when app starts
        self.advanced_btn.setChecked(False)
        self.adv_widget.setVisible(False)

        # Reset status label at the bottom of the screen 
        self.label_5.setStyleSheet("background-color:#cce6ff;")
        self.label_5.setText("🤓Welcome to PDF Shrinker. Let's shrink some PDFs...")

    #Help PushButton Slot
    def show_help(self):
        help_msg=QMessageBox()
        help_msg.setTextFormat(Qt.RichText)
        help_msg.information(MainForm,"Help",'Please visit the project page and read the documentation section at the following URL for more detailed explanations and examples:<br> <a href="https://github.com/chinmayasrathore/pdf-shrinker/">https://github.com/chinmayasrathore/pdf-shrinker/<a> ')
    
    #Gracefully Exit the app on press of the Quit button 
    def exit_app(self):
        self.label_5.setStyleSheet("background-color:#ffccff")
        self.label_5.setText("🤓Thanks for using PDFShrinker. ")
        QTimer.singleShot(1000,QApplication.instance().quit)

    # Action to be taken after successful completion. Some passed on parameters not used. Only for testing. 
    def handle_result(self,returncode,stdout,stderr):
        #Stop the loader gif and hide the label on successful completion 
        self.movie.stop()
        self.loader_label.hide()
        
        #Variables used in success status message
        now = datetime.now()
        output_pdf = self.lineEdit_outputPDF.text().strip()
        file_name = os.path.basename(output_pdf) 

        # Display successful compressed file creation status message 
        self.label_5.setStyleSheet("background-color:#b3ffb3;")
        self.label_5.setText(f"🏆Success! Compressed PDF file {file_name} has been created on {now} .")
        
        # Activate shrink pdf, Clear and Quit pushbuttons and clear output file name 
        self.pushButton_compress.setEnabled(True)
        self.pushButton_clear.setEnabled(True)
        self.pushButton_quit.setEnabled(True)
        self.lineEdit_outputPDF.clear()
        
        # Change the status message to invite a new compression attempt if desired by the user
        # This is done after 10 seconds (10000 milliseconds) of displaying the Success message above
        QTimer.singleShot(10000,self.clear_status) 

    # Handle errors in case the compression attempt was unsuccessful
    # using error message using the captured error in the run method in the Worker class (lines 36-41)   
    def handle_error(self, error_message):

        #Stop the loader GIF and hide the working label 
        self.movie.stop()
        self.loader_label.hide()

        #Update status in label_5 to show Error Message 
        self.label_5.setStyleSheet("background-color:#ffb3b3;")
        self.label_5.setText("😟 There was an Error compressing this PDF file! Output PDF could not be created!")
       
       # Display error message with error details in a message box
        QMessageBox.critical(MainForm, "⚠️ Execution Error !", f"An error has been encountered. Current operation has been terminated. \n\nPossible Cause: Output PDF file is open in a PDF viewer.\nClose the open PDF file or choose a different output PDF file name and try again. \n\n --------\n\nError Diagnostic Information: {error_message} ")
        
        
        # Exit Error handling to the Main App window reactivating the compress button and 
        # refreshing the status in label_5 to permit another compression attempt. 
        self.pushButton_compress.setEnabled(True) # Activate Button
        self.pushButton_clear.setEnabled(True) # Activate Button 
        self.pushButton_quit.setEnabled(True) # Activate Button 

        QTimer.singleShot(1000,self.clear_status) 

          

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    MainForm = QWidget()
    ui = Ui_Form()
    ui.setupUi(MainForm)
    MainForm.show()
    sys.exit(app.exec())

   