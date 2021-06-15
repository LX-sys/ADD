

# 增加号
def serialNum_addPushColor():
    return '''QPushButton{
        background-color: rgb(34,139,34);
        color:rgb(255,255,255);
        border-color: beige;
        border-width: 2px;
        border-radius:10px;
        }
        QPushButton:pressed{
        background-color: rgb(152,251,152);
        color:rgb(0,0,0);
        border-style: inset;
        }'''

# 下拉框
def serialNum_comboBoxColor():
    return '''QComboBox {border:none;
    border-radius:10px;
    background:rgb(255,250,250);
    color:rgb(128,0,0);
    }
    
    '''


# 当前Excel
def excelName_labelColor():
    return '''QLabel{
        color:rgb(139,69,19);
        font-size:14px;
        }
    '''
# 位置
def excelPos_labelColor():
    return '''QLabel{
        color:rgb(139,69,19);
        font-size:14px;
        }
    '''


# 路径标签
def cecelShowPos_labelColor():
    return '''QLabel{
        color:rgb(128,128,128);
        }
    '''
# 刷新
def refresh_pushButtonColor():
    return '''QPushButton{
        background-color: rgb(0,0,0);
        color:rgb(255,255,255);
        border-color: beige;
        border-width: 2px;
        border-radius:10px;
        font-size:18px;
        }
        QPushButton:pressed{
        background-color: rgb(255,255,255);
        color:rgb(0,0,0);
        border-style: inset;
        }'''

# 下载按钮
def pushButtonColor():
    return '''QPushButton{
        background-color: rgb(34,139,34);
        color:rgb(255,255,255);
        border-color: beige;
        border-width: 2px;
        border-radius:20px;
        }
        QPushButton:pressed{
        background-color: rgb(152,251,152);
        color:rgb(0,0,0);
        border-style: inset;
        }'''

# 重置按钮
def init_pushButtonColr():
    return '''QPushButton{
        background-color: rgb(178,34,34);
        color:rgb(255,255,255);
        border-color: beige;
        border-width: 2px;
        border-radius:10px;
        font-size:14px;
        }
        QPushButton:pressed{
        background-color: rgb(224, 0, 0);
        border-style: inset;
        }'''

# 创建按钮
def pushButton_2Colr():
    return '''QPushButton{
        background-color: rgb(128,128,128);
        color:rgb(255,255,255);
        border-color: beige;
        border-width: 2px;
        border-radius:10px;
        font-size:18px;
        }
        QPushButton:pressed{
        background-color: rgb(0,0,0);
        border-style: inset;
        }'''

# 控件左
def hello_pushButtonColor():
    return '''QPushButton{
        background-color: rgb(0,0,205);
        border-style: solid;
        color:rgb(255,255,255);
        border-color: rgb(0,0,205);
        border-width: 2px;
        border-radius:10px;
        }
        QPushButton:pressed{
        background-color: rgb(176,196,222);
        color:rgb(0,0,0);
        border-style: inset;
        }'''
# 控件右
def hello_pushButton_2Color():
    return '''QPushButton{
            background-color: rgb(240,128,128);
            border-style:outset;
            color:rgb(255,255,255);
            border-color: rgb(205,92,92);
            border-width: 2px;
            border-radius:10px;
            }
            QPushButton:pressed{
            background-color: rgb(255,192,203);
            color:rgb(0,0,0);
            border-style: inset;
            }'''
# 自定义控件左
def zdy_push_Color():
    return '''QPushButton{
                background-color: rgb(65,105,225);
                border-style:outset;
                color:rgb(255,255,255);
                border-color: rgb(0,0,205);
                border-width: 2px;
                border-radius:10px;
                }
                QPushButton:pressed{
                background-color: rgb(176,196,222);
                color:rgb(0,0,0);
                border-style: inset;
                }'''
# 自定义控件右
def zdy_push_Color2():
    return '''QPushButton{
                background-color: rgb(255,160,122);
                border-style:outset;
                color:rgb(255,255,255);
                border-color: rgb(205,92,92);
                border-width: 2px;
                border-radius:10px;
                }
                QPushButton:pressed{
                background-color: rgb(255,192,203);
                color:rgb(0,0,0);
                border-style: inset;
                }'''
# 自定义line
def lineColor():
    return '''QLineEdit { 
  background-color: rgb(255,160,122);   /*背景色*/
  color:rgb(0,0,0);    /*前景色*/
  selection-color: blue;  /*文字被选中时的颜色*/ 
  selection-background-color: green; /*文字被选中时的背景色*/ 
}
    '''

# 当前月
def mon_label_EditColor():
    return '''QLabel{
        color:rgb(25,25,112);
        font-size:30px;
        }
    '''
