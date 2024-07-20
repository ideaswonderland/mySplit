from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
import sys
import os
from datetime import datetime
from mySplit import Fatura

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("LETA Fatura")
        self.setWindowIcon(QIcon('logo.png'))
        self.setFixedSize(250,400)

        layout = QVBoxLayout()
        self.setLayout(layout)
        
        btnElk = QPushButton(text="Elektrik")
        btnElk.setFixedSize(230, 60)
        btnElk.clicked.connect(lambda: self.elektrikSender())
        
        btnSu = QPushButton(text="Su")
        btnSu.setFixedSize(230, 60)
        
        btnGaz = QPushButton(text="Doğalgaz")
        btnGaz.setFixedSize(230, 60)
        btnGaz.clicked.connect(lambda: self.gazSender())
        
        btnInt = QPushButton(text="İnternet")
        btnInt.setFixedSize(230, 60)
        btnInt.clicked.connect(lambda: self.internetSender())
        
        btnTel = QPushButton(text="Telefon")
        btnTel.setFixedSize(230, 60)
        btnTel.clicked.connect(lambda: self.telefonSender())
        
        btnBilgi = QPushButton(text="Bilgiler")
        btnBilgi.setFixedSize(230, 60)
        btnBilgi.clicked.connect(lambda: self.bilgiSender())
        
        
        layout.addWidget(btnElk)
        layout.addWidget(btnSu)
        layout.addWidget(btnGaz)
        layout.addWidget(btnInt)
        layout.addWidget(btnTel)
        layout.addWidget(btnBilgi)
           
    def elektrikSender(self):
        self.w = Elektrik()
        self.w.show()
        self.close()
        
    def internetSender(self):
        self.w = İnternet()
        self.w.show()
        self.close()
        
    def telefonSender(self):
        self.w = Telefon()
        self.w.show()
        self.close()
        
    def gazSender(self):
        self.w = Doğalgaz()
        self.w.show()
        self.close()
        
    def bilgiSender(self):
        self.w = Bilgi()
        self.w.show()
        self.close()
        
class Elektrik(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("LETA Fatura")
        self.setWindowIcon(QIcon('logo.png'))
        self.setFixedSize(250,180)

        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.label1 = QLabel('Ay:',self) 
        self.label1.setGeometry(15, 0, 110, 40)
        self.label2 = QLabel('Yıl:',self)
        self.label2.setGeometry(125, 0, 110, 40)
        
        self.spinAy = QSpinBox(self)
        self.spinAy.setGeometry(10, 30, 110, 40)
        self.spinAy.setMinimum(1)
        self.spinAy.setMaximum(12)
        self.spinAy.setValue(datetime.now().month-1)
        
        self.spinYıl = QSpinBox(self)
        self.spinYıl.setGeometry(120, 30, 110, 40)
        self.spinYıl.setMinimum(2017)
        self.spinYıl.setMaximum(2027)
        self.spinYıl.setValue(datetime.now().year)
          
        btnHesapla = QPushButton(text='Hesapla', parent=self)
        btnHesapla.setFixedSize(230, 30)
        btnHesapla.setStyleSheet("background-color : #57f752")
        btnHesapla.move(10,80)
        btnHesapla.clicked.connect(lambda: self.hesapla())        
        
        btnGeri = QPushButton(text='Geri Dön', parent=self)
        btnGeri.setFixedSize(230, 30)
        btnGeri.setStyleSheet("background-color : #c8d3fa")
        btnGeri.move(10,130)
        btnGeri.clicked.connect(lambda: self.the_back_button_was_clicked())

    
    def the_back_button_was_clicked(self):
        self.w = MainWindow()
        self.w.show()
        self.close()      
        
    def hesapla(self):
        Fatura(self.spinAy.value(),self.spinYıl.value()).Elektrik()
        self.showDialog()


    def showDialog(self):
        msgBox = QMessageBox(self)
        msgBox.setText("Hesaplama tamamlandı!")
        msgBox.setWindowTitle("LETA Fatura")
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setIcon(QMessageBox.Icon.Information)

        returnValue = msgBox.exec()

        
        
        
        
class İnternet(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("LETA Fatura")
        self.setWindowIcon(QIcon('logo.png'))
        self.setFixedSize(250,180)

        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.label1 = QLabel('Ay:',self) 
        self.label1.setGeometry(15, 0, 110, 40)
        self.label2 = QLabel('Yıl:',self)
        self.label2.setGeometry(125, 0, 110, 40)
        
        self.spinAy = QSpinBox(self)
        self.spinAy.setGeometry(10, 30, 110, 40)
        self.spinAy.setMinimum(1)
        self.spinAy.setMaximum(12)
        self.spinAy.setValue(datetime.now().month-1)
        
        self.spinYıl = QSpinBox(self)
        self.spinYıl.setGeometry(120, 30, 110, 40)
        self.spinYıl.setMinimum(2017)
        self.spinYıl.setMaximum(2027)
        self.spinYıl.setValue(datetime.now().year)
          
        btnHesapla = QPushButton(text='Hesapla', parent=self)
        btnHesapla.setFixedSize(230, 30)
        btnHesapla.setStyleSheet("background-color : #57f752")
        btnHesapla.move(10,80)
        btnHesapla.clicked.connect(lambda: self.hesapla())        
        
        btnGeri = QPushButton(text='Geri Dön', parent=self)
        btnGeri.setFixedSize(230, 30)
        btnGeri.setStyleSheet("background-color : #c8d3fa")
        btnGeri.move(10,130)
        btnGeri.clicked.connect(lambda: self.the_back_button_was_clicked())

    
    def the_back_button_was_clicked(self):
        self.w = MainWindow()
        self.w.show()
        self.close()      
        
    def hesapla(self):
        Fatura(self.spinAy.value(),self.spinYıl.value()).İnternet()
        self.showDialog()


    def showDialog(self):
        msgBox = QMessageBox(self)
        msgBox.setText("Hesaplama tamamlandı!")
        msgBox.setWindowTitle("LETA Fatura")
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setIcon(QMessageBox.Icon.Information)

        returnValue = msgBox.exec()        
        
        
class Telefon(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("LETA Fatura")
        self.setWindowIcon(QIcon('logo.png'))
        self.setFixedSize(250,180)

        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.label1 = QLabel('Ay:',self) 
        self.label1.setGeometry(15, 0, 110, 40)
        self.label2 = QLabel('Yıl:',self)
        self.label2.setGeometry(125, 0, 110, 40)
        
        self.spinAy = QSpinBox(self)
        self.spinAy.setGeometry(10, 30, 110, 40)
        self.spinAy.setMinimum(1)
        self.spinAy.setMaximum(12)
        self.spinAy.setValue(datetime.now().month-1)
        
        self.spinYıl = QSpinBox(self)
        self.spinYıl.setGeometry(120, 30, 110, 40)
        self.spinYıl.setMinimum(2017)
        self.spinYıl.setMaximum(2027)
        self.spinYıl.setValue(datetime.now().year)
          
        btnHesapla = QPushButton(text='Hesapla', parent=self)
        btnHesapla.setFixedSize(230, 30)
        btnHesapla.setStyleSheet("background-color : #57f752")
        btnHesapla.move(10,80)
        btnHesapla.clicked.connect(lambda: self.hesapla())        
        
        btnGeri = QPushButton(text='Geri Dön', parent=self)
        btnGeri.setFixedSize(230, 30)
        btnGeri.setStyleSheet("background-color : #c8d3fa")
        btnGeri.move(10,130)
        btnGeri.clicked.connect(lambda: self.the_back_button_was_clicked())

    
    def the_back_button_was_clicked(self):
        self.w = MainWindow()
        self.w.show()
        self.close()      
        
    def hesapla(self):
        Fatura(self.spinAy.value(),self.spinYıl.value()).Telefon()
        self.showDialog()


    def showDialog(self):
        msgBox = QMessageBox(self)
        msgBox.setText("Hesaplama tamamlandı!")
        msgBox.setWindowTitle("LETA Fatura")
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setIcon(QMessageBox.Icon.Information)

        returnValue = msgBox.exec()

class Doğalgaz(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("LETA Fatura")
        self.setWindowIcon(QIcon('logo.png'))
        self.setFixedSize(250,180)

        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.label1 = QLabel('Ay:',self) 
        self.label1.setGeometry(15, 0, 110, 40)
        self.label2 = QLabel('Yıl:',self)
        self.label2.setGeometry(125, 0, 110, 40)
        
        self.spinAy = QSpinBox(self)
        self.spinAy.setGeometry(10, 30, 110, 40)
        self.spinAy.setMinimum(1)
        self.spinAy.setMaximum(12)
        self.spinAy.setValue(datetime.now().month-1)
        
        self.spinYıl = QSpinBox(self)
        self.spinYıl.setGeometry(120, 30, 110, 40)
        self.spinYıl.setMinimum(2017)
        self.spinYıl.setMaximum(2027)
        self.spinYıl.setValue(datetime.now().year)
          
        btnHesapla = QPushButton(text='Hesapla', parent=self)
        btnHesapla.setFixedSize(230, 30)
        btnHesapla.setStyleSheet("background-color : #57f752")
        btnHesapla.move(10,80)
        btnHesapla.clicked.connect(lambda: self.hesapla())        
        
        btnGeri = QPushButton(text='Geri Dön', parent=self)
        btnGeri.setFixedSize(230, 30)
        btnGeri.setStyleSheet("background-color : #c8d3fa")
        btnGeri.move(10,130)
        btnGeri.clicked.connect(lambda: self.the_back_button_was_clicked())

    
    def the_back_button_was_clicked(self):
        self.w = MainWindow()
        self.w.show()
        self.close()      
        
    def hesapla(self):
        Fatura(self.spinAy.value(),self.spinYıl.value()).Doğalgaz()
        self.showDialog()


    def showDialog(self):
        msgBox = QMessageBox(self)
        msgBox.setText("Hesaplama tamamlandı!")
        msgBox.setWindowTitle("LETA Fatura")
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setIcon(QMessageBox.Icon.Information)

        returnValue = msgBox.exec()


class Bilgi(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("LETA Fatura")
        self.setWindowIcon(QIcon('logo.png'))
        self.setFixedSize(250,250)

        layout = QVBoxLayout()
        self.setLayout(layout)
        
          
        """btnHesapla = QPushButton(text='Hesapla', parent=self)
        btnHesapla.setFixedSize(230, 30)
        btnHesapla.setStyleSheet("background-color : #57f752")
        btnHesapla.move(10,80)
        btnHesapla.clicked.connect(lambda: self.hesapla())"""        
        
        btnGeri = QPushButton(text='Geri Dön', parent=self)
        btnGeri.setFixedSize(230, 30)
        btnGeri.setStyleSheet("background-color : #c8d3fa")
        btnGeri.move(10,150)
        btnGeri.clicked.connect(lambda: self.the_back_button_was_clicked())

        btnDosya = QPushButton(text='Dosya Yolu', parent=self)
        btnDosya.setFixedSize(230, 30)
        btnDosya.setStyleSheet("background-color : #ffbf80")
        btnDosya.move(10,10)
        btnDosya.clicked.connect(lambda: self.dosyaYolu()) 
        
        btnFirma = QPushButton(text='Firma Bilgileri', parent=self)
        btnFirma.setFixedSize(230, 30)
        btnFirma.setStyleSheet("background-color : #ffbf80")
        btnFirma.move(10,45)
        btnFirma.clicked.connect(lambda: self.firmaBilgileri())  
        
        btnİmza = QPushButton(text='İmza Bilgileri', parent=self)
        btnİmza.setFixedSize(230, 30)
        btnİmza.setStyleSheet("background-color : #ffbf80")
        btnİmza.move(10,80)
        btnİmza.clicked.connect(lambda: self.imzaBilgileri())     
        
        btnKurum = QPushButton(text='Kurum Bilgileri', parent=self)
        btnKurum.setFixedSize(230, 30)
        btnKurum.setStyleSheet("background-color : #ffbf80")
        btnKurum.move(10,115)
        btnKurum.clicked.connect(lambda: self.kurumBilgileri())                      
        
    
    def the_back_button_was_clicked(self):
        self.w = MainWindow()
        self.w.show()
        self.close()      
                
    def dosyaYolu(self):
        response = QFileDialog.getExistingDirectory(
            self,
            caption='Fatura Dosyanızı Seçiniz!'
        )
        f = open("dosyaYolu.txt", "w")
        f.write(response)
        f.close()
        return
    
    def firmaBilgileri(self):
        os.system("start EXCEL.EXE firma.xlsx")
        return
    
    def imzaBilgileri(self):
        os.system("start EXCEL.EXE imza.xlsx")
        return
    
    def kurumBilgileri(self):
        os.system("start EXCEL.EXE kurum.xlsx")
        return
        
        
        
    
app = QApplication(sys.argv)

w = MainWindow()
w.show()

sys.exit(app.exec())