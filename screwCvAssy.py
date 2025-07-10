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

parts=['screwCvAssy',]

S_Dia=['300','350','400','450','500',]


# 画面を並べて表示する
class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(250, 520)
        Dialog.move(1000, 0)


        #図形
        self.label_6 = QtGui.QLabel(Dialog)
        self.label_6.setGeometry(QtCore.QRect(35, 230, 200, 200))
        self.label_6.setText("")
        
        base=os.path.dirname(os.path.abspath(__file__))
        joined_path = os.path.join(base, "screwCv.png")
        self.label_6.setPixmap(QtGui.QPixmap(joined_path))
        
        self.label_6.setAlignment(QtCore.Qt.AlignTop)
        self.label_6.setObjectName("label_6")

        #screwDia
        self.label_Dia = QtGui.QLabel('screwDia',Dialog)
        self.label_Dia.setGeometry(QtCore.QRect(10, 13, 130, 22))
        self.comboBox_Dia = QtGui.QComboBox(Dialog)
        self.comboBox_Dia.setGeometry(QtCore.QRect(150, 13, 50, 22))
        #self.comboBox_Dia.listIndex=11
        
        #screwCv Length
        self.label_L = QtGui.QLabel('screwCv Length',Dialog)
        self.label_L.setGeometry(QtCore.QRect(10, 63, 100, 22))
        self.le_L = QtGui.QLineEdit(Dialog)
        self.le_L.setGeometry(QtCore.QRect(150, 63, 50, 20))
        self.le_L.setAlignment(QtCore.Qt.AlignCenter)

        #作成
        self.pushButton = QtGui.QPushButton('Create',Dialog)
        self.pushButton.setGeometry(QtCore.QRect(40, 95, 50, 22))
        #インポート
        self.pushButton4 = QtGui.QPushButton('Import',Dialog)
        self.pushButton4.setGeometry(QtCore.QRect(40, 120, 180, 22))
        #更新
        self.pushButton2 = QtGui.QPushButton('upDate',Dialog)
        self.pushButton2.setGeometry(QtCore.QRect(130, 95, 50, 22))

        #spinBox
        #self.label_spin=QtGui.QLabel('Animation',Dialog)
        #self.label_spin.setGeometry(QtCore.QRect(10, 145, 150, 22))
        #spinBox
        #self.spinBox=QtGui.QSpinBox(Dialog)
        #self.spinBox.setGeometry(100, 145, 50, 50)
        #self.spinBox.setMinimum(0.0)  # 最小値を0.0に設定
        #self.spinBox.setMaximum(100.0)  # 最大値を200.0に設定
        #self.spinBox.setValue(100.0)
        #self.spinBox.setAlignment(QtCore.Qt.AlignCenter)
        ##base
        #self.label_kiten = QtGui.QLabel('Base',Dialog)
        #self.label_kiten.setGeometry(QtCore.QRect(10, 195, 100, 22))
        #self.le_kiten = QtGui.QLineEdit('100',Dialog)
        #self.le_kiten.setGeometry(QtCore.QRect(100, 195, 120, 20))
        #self.le_kiten.setAlignment(QtCore.Qt.AlignCenter)

        #ローラーチェン
        #self.pushButton4 = QtGui.QPushButton('ローラーチェン',Dialog)
        #self.pushButton4.setGeometry(QtCore.QRect(10, 640, 100, 22))
        #sprocket調節
        #self.spinBox2=QtGui.QSpinBox(Dialog)
        #self.spinBox2.setGeometry(170, 145, 50, 50)
        #self.spinBox2.setMinimum(0.0)  # 最小値を0.0に設定
        #self.spinBox2.setMaximum(100.0)  # 最大値を100.0に設定
        #self.spinBox2.setValue(0.0)
        #self.spinBox2.setAlignment(QtCore.Qt.AlignCenter)

        #質量計算
        self.pushButton_m = QtGui.QPushButton('massCulculation',Dialog)
        self.pushButton_m.setGeometry(QtCore.QRect(10, 410, 100, 23))
        self.pushButton_m.setObjectName("pushButton") 
        #質量集計
        self.pushButton_m20 = QtGui.QPushButton('massTally_csv',Dialog)
        self.pushButton_m20.setGeometry(QtCore.QRect(110, 410, 130, 23))
        self.pushButton_m2 = QtGui.QPushButton('massTally_SpreadSheet',Dialog)
        self.pushButton_m2.setGeometry(QtCore.QRect(10, 435, 130, 23))
        #質量入力
        self.pushButton_m3 = QtGui.QPushButton('massImput[kg]',Dialog)
        self.pushButton_m3.setGeometry(QtCore.QRect(10, 460, 100, 23))
        self.pushButton_m3.setObjectName("pushButton")  
        self.le_mass = QtGui.QLineEdit(Dialog)
        self.le_mass.setGeometry(QtCore.QRect(110, 460, 50, 20))
        self.le_mass.setAlignment(QtCore.Qt.AlignCenter)  
        self.le_mass.setText('10.0')
        #密度
        self.lbl_gr = QtGui.QLabel('SpecificGravity',Dialog)
        self.lbl_gr.setGeometry(QtCore.QRect(10, 485, 80, 12))
        self.le_gr = QtGui.QLineEdit(Dialog)
        self.le_gr.setGeometry(QtCore.QRect(110, 485, 50, 20))
        self.le_gr.setAlignment(QtCore.Qt.AlignCenter)  
        self.le_gr.setText('7.85')

        self.comboBox_Dia.addItems(S_Dia)
        #self.spinBox.valueChanged[int].connect(self.spinMove)
        #self.le_kiten.textChanged.connect(self.update_kiten)
        #self.le_L.setText('3500')

        QtCore.QObject.connect(self.pushButton4, QtCore.SIGNAL("pressed()"), self.onImport)
        QtCore.QObject.connect(self.pushButton2, QtCore.SIGNAL("pressed()"), self.update)
        self.retranslateUi(Dialog)
        QtCore.QObject.connect(self.pushButton, QtCore.SIGNAL("pressed()"), self.create)
        QtCore.QObject.connect(self.pushButton_m, QtCore.SIGNAL("pressed()"), self.massCulc)
        QtCore.QObject.connect(self.pushButton_m2, QtCore.SIGNAL("pressed()"), self.massTally)
        QtCore.QObject.connect(self.pushButton_m20, QtCore.SIGNAL("pressed()"), self.massTally2)
        QtCore.QObject.connect(self.pushButton_m3, QtCore.SIGNAL("pressed()"), self.massImput)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QtGui.QApplication.translate("Dialog", "Gland packing Assy", None))
        
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

    def massImput(self):
         # 選択したオブジェクトを取得する
        c00 = Gui.Selection.getSelection()
        if c00:
            obj = c00[0]
        label='mass[kg]'
        g=float(self.le_mass.text())
        try:
            obj.addProperty("App::PropertyFloat", "mass",label)
            obj.mass=g
        except:
            obj.mass=g
         
    def massCulc(self):
        c00 = Gui.Selection.getSelection()
        if c00:
            obj = c00[0]
        label='mass[kg]'
        g0=float(self.le_gr.text())
        g=obj.Shape.Volume*g0*1000/10**9  
        try:
            obj.addProperty("App::PropertyFloat", "mass",label)
            obj.mass=g
        except:
            pass

    def massTally2(self):#csv
        doc = App.ActiveDocument
        objects = doc.Objects
        mass_list = []
        for obj in objects:
            if Gui.ActiveDocument.getObject(obj.Name).Visibility:
                if obj.isDerivedFrom("Part::Feature"):
                    if hasattr(obj, "mass"):
                        try:
                            mass_list.append([obj.Label, obj.dia,'1', obj.mass])
                        except:
                            mass_list.append([obj.Label, '','1', obj.mass])    

                else:
                     pass
        doc_path = doc.FileName
        csv_filename = os.path.splitext(os.path.basename(doc_path))[0] + "_counts_and_masses.csv"
        csv_path = os.path.join(os.path.dirname(doc_path), csv_filename)
        with open(csv_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Name",'Standard','Count', "Mass[kg]"])
            writer.writerows(mass_list) 
    def massTally(self):#spreadsheet
        doc = App.ActiveDocument
        # 新しいスプレッドシートを作成
        spreadsheet = doc.addObject("Spreadsheet::Sheet", "PartList")
        spreadsheet.Label = "Parts List"
        
        # ヘッダー行を記入
        headers = ['No',"Name",'Standard', 'Count','Unit[kg]','Mass[kg]']
        for header in enumerate(headers):
            spreadsheet.set(f"A{1}", headers[0])
            spreadsheet.set(f"B{1}", headers[1])
            spreadsheet.set(f"C{1}", headers[2])
            spreadsheet.set(f"D{1}", headers[3])
            spreadsheet.set(f"E{1}", headers[4])
            spreadsheet.set(f"F{1}", headers[5])
        # パーツを列挙して情報を書き込む
        row = 2
        i=1
        s=0
        for i,obj in enumerate(doc.Objects):
            if obj.Label=='本体' or obj.Label=='本体 (mirrored)' or obj.Label[:7]=='Channel' or obj.Label[:5]=='Angle' \
                or obj.Label[:6]=='Square' or obj.Label[:7]=='Extrude' or obj.Label[:6]=='Fusion' or obj.Label[:6]=='Corner' \
                    or obj.Label[:5]=='basic' or obj.Label[:4]=='Edge' or obj.Label[:3]=='hub' or obj.Label[:7]=='_8_tube'\
                        or obj.Label[:5]=='plate' or obj.Label[:6]=='keyway' or obj.Label[:4]=='tube':
                pass        
            else:  
                try:
                    spreadsheet.set(f"E{row}", f"{obj.mass:.2f}")  # Unit
                    s=obj.mass+s
                    if hasattr(obj, "Shape") and obj.Shape.Volume > 0:
                        try:
                            spreadsheet.set(f"A{row}", str(row-1))  # No
                            spreadsheet.set(f"B{row}", obj.Label)   #Name
                            try:
                                spreadsheet.set(f"C{row}", obj.dia)
                            except:
                                #spreadsheet.set(f"C{row}", obj.standard)
                                pass
                            if obj.Label[:7]=='Angular':
                                n=2
                            else:
                                n=1    
                            spreadsheet.set(f"D{row}", str(n))   # count
                            g=round(obj.mass*n,2)
                            spreadsheet.set(f"F{row}", str(g))   # g

                    
                            row += 1
                        except:
                            pass    
                except:
                    pass
                spreadsheet.set(f'F{row}',str(s))
        App.ActiveDocument.recompute()
        Gui.activeDocument().activeView().viewAxometric()
    
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