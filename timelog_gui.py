# :coding: utf-8

import sys
import re
import os
import datetime as dt
from pprint import pprint

from shotgun_api3 import Shotgun

from collections import OrderedDict


from PyQt5.QtCore import *
from PyQt5.QtGui  import *
from PyQt5.QtWidgets  import *

import timelogs

sg = Shotgun(
    'https://west.shotgunstudio.com',
    'ami_main',
    'iuqeijx^vfdaZcphhjxqxdp4z'
)


class Timelog_GUI( QDialog ):
    def __init__( self ):
        super( Timelog_GUI, self ).__init__() 
        
        self.resize( 800, 400 )
        self.setWindowTitle( 'Timelog Extractor' )           
        
        dept_lb = QLabel( '[ Department ]' )
        user_lb = QLabel( '[ User ]' )
        self.dept_cb = QComboBox()
        self.user_cb = QComboBox()
        self.dept_cb.setFixedHeight( 35 )
        self.user_cb.setFixedHeight( 35 )

        dept_lay = QVBoxLayout()
        dept_lay.addWidget( dept_lb )
        dept_lay.addWidget( self.dept_cb )
        user_lay = QVBoxLayout()
        user_lay.addWidget( user_lb )
        user_lay.addWidget( self.user_cb )

        cb_lay = QHBoxLayout()
        cb_lay.addLayout( dept_lay )
        cb_lay.addLayout( user_lay )

        self.start_cal = QCalendarWidget()
        self.end_cal   = QCalendarWidget()
        cal_lay = QHBoxLayout()
        cal_lay.addWidget( self.start_cal )
        cal_lay.addWidget( self.end_cal   )


        self.export_btn = QPushButton( 'Export Excel' )
        self.cancel_btn  = QPushButton('Cancel')
        self.export_btn.setFixedHeight( 35 )
        self.cancel_btn.setFixedHeight( 35 )
        btn_lay = QHBoxLayout()
        btn_lay.addWidget( self.export_btn )
        btn_lay.addWidget( self.cancel_btn )

        main_lay = QVBoxLayout()
        main_lay.addLayout( cb_lay )
        main_lay.addLayout( cal_lay )
        sp = QSpacerItem( 10,20 )
        main_lay.addSpacerItem( sp )
        main_lay.addLayout( btn_lay )
        self.setLayout( main_lay )
        

        today = QDate.currentDate()
        self.end_cal.setSelectedDate( today )
        self.start_cal.setSelectedDate( today.addYears(-1) )
        
        self.cancel_btn.clicked.connect( self.close )
        self.dept_cb.currentIndexChanged.connect( self.init_user ) 
        self.export_btn.clicked.connect( self.export_excel )

        self.init_dept()
        #self.init_user()


    def init_dept( self ):
        result = sg.find(
            'HumanUser',
            [
                ['sg_status_list', 'is', 'act'],
                ['department.Department.name', 'is_not', '']
                
            ],
            [ 'firstname', 'department.Department.name',]
        )

        dict = {}
        
        for x in result:
            if x['department.Department.name'] not in ['HQ','ManagementSupport','vendor', '외주사',]:
                if x['department.Department.name'] in list( dict.keys() ):
                    dict[ x['department.Department.name'] ].append( x['firstname'] )
                else:
                    dict[ x['department.Department.name'] ] = [ x['firstname'] ]


        self.dict = dict

        self.dept_cb.addItems( sorted( [ x for x in  list( self.dict.keys() )]) )
        self.dept_cb.setCurrentIndex( 0 )
        

    def init_user( self ):
        dept = str( self.dept_cb.currentText() )
        #print( dept )
        self.user_cb.clear()
        for x in self.dict:
            if x == dept :
                self.user_cb.addItems(  sorted( [ n for n in self.dict[x] ] ) ) 
                #print( sorted( [ n for n in self.dict[x] ] ) )
                break

    def export_excel( self ):

        start_date = self.start_cal.selectedDate()          
        end_date   = self.end_cal.selectedDate()          
        user       = str( self.user_cb.currentText() )

        save_path = QFileDialog.getSaveFileName( 
                self, 'Save file', 
                os.path.expanduser( '~' ) + os.sep + user + '_' + dt.datetime.today().strftime( '%y%m%d' ),
                '*.xlsx' 
        )
        
        #pprint( save_path )
        if os.path.splitext( save_path[0] )[1] != '.xlsx':
            save_path = save_path[0] + '.xlsx'
            
            
        self.setWindowTitle( 'Eporting Excel...' )
        timelogs.main( 
            user,
            start_date.year(), start_date.month(), start_date.day(),
            end_date.year(), end_date.month(), end_date.day(),
            save_path
        )
        self.setWindowTitle( 'Timelog Extractor' )           
        QMessageBox.information( self,'Done!!',  'Finished to export excel')
        
        #self.close()



if __name__ == '__main__':

    app = QApplication( sys.argv )
    mainWin = Timelog_GUI()
    mainWin.show()
    sys.exit( app.exec_() )

