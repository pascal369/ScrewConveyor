# -*- coding: utf-8 -*-
import os
import sys
import Import
import Spreadsheet
import DraftVecUtils
import Sketcher
import PartDesign
import FreeCAD as App
import FreeCADGui as Gui
from PySide import QtGui
from PySide import QtUiTools
from PySide import QtCore
import FreeCAD
import csv
parts=['screwCvAssy',]

S_Dia=['300','350','400','450','500',]


# 画面を並べて表示する
class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(250, 300)
        Dialog.move(1000, 0)

        #screwDia
        self.label_Dia = QtGui.QLabel('screwDia',Dialog)
        self.label_Dia.setGeometry(QtCore.QRect(10, 13, 130, 22))
        self.label_Dia.setStyleSheet("color: black;")
        self.comboBox_Dia = QtGui.QComboBox(Dialog)
        self.comboBox_Dia.setGeometry(QtCore.QRect(150, 8, 50, 22))

        #screwCv Length
        self.label_L = QtGui.QLabel('screwCv Length',Dialog)
        self.label_L.setGeometry(QtCore.QRect(10, 35, 100, 22))
        self.label_L.setStyleSheet("color: black;")
        self.le_L = QtGui.QLineEdit(Dialog)
        self.le_L.setGeometry(QtCore.QRect(150, 35, 50, 22))
        self.le_L.setAlignment(QtCore.Qt.AlignCenter)

        #作成
        self.pushButton = QtGui.QPushButton('Create',Dialog)
        self.pushButton.setGeometry(QtCore.QRect(35, 60, 45, 22))
        #更新
        self.pushButton2 = QtGui.QPushButton('upDate',Dialog)
        self.pushButton2.setGeometry(QtCore.QRect(130, 60, 50, 22))
        self.comboBox_Dia.addItems(S_Dia)
        #インポート
        self.pushButton4 = QtGui.QPushButton('Import',Dialog)
        self.pushButton4.setGeometry(QtCore.QRect(35, 85, 185, 22))
        #図形
        self.label_6 = QtGui.QLabel(Dialog)
        self.label_6.setGeometry(QtCore.QRect(35, 120, 200, 200))
        self.label_6.setText("")
        base=os.path.dirname(os.path.abspath(__file__))
        joined_path = os.path.join(base, "screwCv.png")
        self.label_6.setPixmap(QtGui.QPixmap(joined_path))
        self.label_6.setAlignment(QtCore.Qt.AlignTop)
        self.label_6.setObjectName("label_6")

        QtCore.QObject.connect(self.pushButton4, QtCore.SIGNAL("pressed()"), self.onImport)
        QtCore.QObject.connect(self.pushButton2, QtCore.SIGNAL("pressed()"), self.update)
        self.retranslateUi(Dialog)
        QtCore.QObject.connect(self.pushButton, QtCore.SIGNAL("pressed()"), self.create)
        
        QtCore.QMetaObject.connectSlotsByName(Dialog)
    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QtGui.QApplication.translate("Dialog", "Screw Conveyor Assy", None))
        
    def onImport(self):
        global spro1
        global spro2 
        global SP1
        global SP2
        global chainPath
        global shtAssy
        global shtCvAssy
        global shtSpro1
        global shtSpro2
        global shtLink
        global scwAssy
        global Angular
        doc = FreeCAD.activeDocument()
        if doc:
            group_names = []
            for obj in doc.Objects:
                print(obj.Label)
                if obj.Label=='spro1':
                    spro1=obj
                elif obj.Label=='spro2':
                    spro2=obj 
                elif obj.Label=='SP1':
                    SP1=obj
                elif obj.Label=='SP2':
                    SP2=obj        
                elif obj.Label=='chainPath':
                    chainPath=obj
                elif obj.Label=='shtAssy':
                    shtAssy=obj 
                elif obj.Label=='shtSpro1':
                    shtSpro1=obj 
                elif obj.Label=='shtSpro1':
                    shtSpro2=obj 
                elif obj.Label=='shtLink':
                    shtLink=obj   
                elif obj.Label == "shtCvAssy":
                         shtCvAssy = obj
                elif obj.Label == "scwAssy":
                         scwAssy = obj   
                elif obj.Label[:7]=='Angular':
                         Angular=obj               


                         self.comboBox_Dia.setCurrentText(shtCvAssy.getContents('scwDia'))  
                         self.le_L.setText(shtCvAssy.getContents('L0'))            

    def update(self):

         L0=self.le_L.text()
         scwDia=self.comboBox_Dia.currentText()
         shtCvAssy.set('L0',str(L0))
         shtCvAssy.set('scwDia',scwDia)

         App.ActiveDocument.recompute()

    def create(self): 

         fname='screwCvAssy.FCStd'
         base=os.path.dirname(os.path.abspath(__file__))
         joined_path = os.path.join(base, fname) 

         try:
            Gui.ActiveDocument.mergeProject(joined_path)
         except:
            doc=App.newDocument()
            Gui.ActiveDocument.mergeProject(joined_path)

    def spinMove(self):
         N1=shtAssy.getContents('Teeth1')
         N2=shtAssy.getContents('Teeth2')
         #try:
         Pitch=shtAssy.getContents('Pitch')
         A=self.spinBox.value()
         beta1=360/float(N1)
         beta2=360/float(N2)
         x=10
         #print(N1,N2,beta1)
         spro1.Placement.Rotation=App.Rotation(App.Vector(0,1,0),A*beta1/x)
         spro2.Placement.Rotation=App.Rotation(App.Vector(0,1,0),A*beta2/x)
         scwAssy.Placement.Rotation=App.Rotation(App.Vector(1,0,0),A*beta2/x)
         #print(A,Pitch,x)
         if A==0:
             return
         self.le_kiten.setText(str(round(A*float(Pitch)/x,3)))

    def update_kiten(self):
        kiten=self.le_kiten.text()
        try:
             shtAssy.set('kiten',kiten)
             print(kiten)
             App.ActiveDocument.recompute() 
        except:
             pass
    def setIchi(self):
         selection = Gui.Selection.getSelection()
         if selection:
                 selected_object = selection[0]
                 if selected_object.TypeId == "App::Part":
                     parts_group = selected_object
                     for obj in parts_group.Group:
                         if obj.TypeId == "Spreadsheet::Sheet":
                             spreadsheet = obj
#
                             try:
                                 if selected_object.Label=='spro1' :
                                     A0=float(self.spinBox2.value())*0.5
                                     SP1.Placement.Rotation=App.Rotation(App.Vector(0,1,0),A0)
                                     spreadsheet.set('spr1',str(A0))
                                     App.ActiveDocument.recompute() 
                                 elif selected_object.Label=='spro2' :
                                     A1=float(self.spinBox2.value())*0.5
                                     SP2.Placement.Rotation=App.Rotation(App.Vector(0,1,0),A1)  
                                     spreadsheet.set('spr2',str(A1))  
                                     App.ActiveDocument.recompute() 
                             except:
                                return                

         
         
class main():
        d = QtGui.QWidget()
        d.ui = Ui_Dialog()
        d.ui.setupUi(d)
        d.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        d.show()  
        #script_window = Gui.getMainWindow().findChild(QtGui.QDialog, 'd')
        #script_window.setWindowFlags(script_window.windowFlags() & ~QtCore.Qt.WindowCloseButtonHint)            