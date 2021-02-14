import sys
import openpyxl
import io
import string
import copy
import os
from PIL import Image as PImage
from openpyxl.drawing.image import Image
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QFileDialog, QVBoxLayout, QLabel, QHBoxLayout


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.select_excel = None
        self.initUI()

    def initUI(self):
        self.excel_name = QLabel()
        select_excel_button = QPushButton('엑셀 불러오기')
        garo = QLabel('가로')
        sero = QLabel('세로')
        self.width = QLineEdit()
        self.height = QLineEdit()
        done_button = QPushButton('작업 시작')

        hbox = QHBoxLayout()
        hbox.addWidget(garo)
        hbox.addWidget(self.width)
        hbox.addWidget(sero)
        hbox.addWidget(self.height)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addWidget(self.excel_name)
        vbox.addWidget(select_excel_button)
        vbox.addLayout(hbox)
        vbox.addWidget(done_button)
        vbox.addStretch(1)

        self.setLayout(vbox)
        
        select_excel_button.clicked.connect(self.select_excel_clicked)
        done_button.clicked.connect(self.done_button_clicked)

        self.setWindowTitle('xpa')
        self.resize(400, 200)


        self.show()

    def done_button_clicked(self):
        resolution = 37.7952755906

        if self.width.text() == '' or self.height.text() == '':
            print('아무것도 입력 안 함')
        else:
            # 핵심 코드 시작
            wb = openpyxl.load_workbook(self.select_excel)

            img_width = float(self.width.text())
            img_height = float(self.height.text())

            sheet_name_list = wb.sheetnames

            # 시트를 순회한다.
            for sheet_name in sheet_name_list:
                print(f'sheet name: {sheet_name}')
                sheet = wb[sheet_name]
                print(f'{sheet_name} images: {sheet._images}')

                print()

                _images = {}

                # 시트에 들어있는 이미지들을 복사해서 리스트로 만든다.
                # 이미지를 삭제할 때, 리스트가 변경되어 for 문에 영향을 주는것을 방지하기 위함
                copy_sheet_images = sheet._images[:]

                # 어느 셀에 무슨 사진이 있는지 쉽게 알 수 있게 {key: 셀, value: 이미지 정보} 인 딕셔너리를 만듬
                for image in copy_sheet_images:
                    row = image.anchor._from.row + 1
                    col = string.ascii_uppercase[image.anchor._from.col]
                    cell = f'{col}{row}'

                    _images[cell] = image._data
                    print(f'_images[cell]: {_images[cell]}')

                print()

                # 시트의 이미지들을 순회하면서
                for image in copy_sheet_images:
                    print(f'image: {image}')
                    row = image.anchor._from.row + 1
                    col = string.ascii_uppercase[image.anchor._from.col]
                    cell = f'{col}{row}'

                    # 사진들이 임의로 저장될 폴더를 하나 만들어줌. ( Images )
                    file_name = os.path.join(os.getcwd(), 'Images')
                    if not os.path.exists(file_name):
                        os.mkdir(file_name)
                    file_name = os.path.join(file_name, f'{cell}.jpg')

                    print(f'cell: {cell}')
                    print(f'file_name: {file_name}')

                    # 핵심코드
                    # _images 딕셔너리의 key: cell 의 value 는 image.data 가 들어있다.
                    # image.data() 는 이미지 정보를 리턴한다.
                    # 즉, 함수가 들어있기 때문에 ()를 붙여서 이미지 정보를 io.BytesIO 로 불러온다.
                    # 그 정보를 PIL 을 이용해 저장해준다.
                    img = io.BytesIO(_images[cell]())
                    img = PImage.open(img)
                    img.save(file_name)

                    # PIL 을 이용해 저장한 사진을 openpyxl 의 Image 를 이용해 불러온다.
                    # width 와 height 를 조정해 이미지 크기를 맞춰주고
                    # cell 위치에 사진을 집어 넣는다.
                    insert_img = Image(file_name)
                    insert_img.width = img_width * resolution
                    insert_img.height = img_height * resolution
                    sheet.add_image(insert_img, cell)

                    # 정렬이 끝난 사진은 삭제를 한다.
                    sheet._images.remove(image)

                    print()

            wb.save('out.xlsx')

    def select_excel_clicked(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Open File', './')

        if fname:
            self.select_excel = fname
            self.excel_name.setText(fname)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())