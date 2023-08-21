#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "xlsxdocument.h"





MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_actionOpen_Excel_File_triggered()
{
    QAxObject* excelApp = new QAxObject("Excel.Application");
    excelApp->setProperty("Visible", true);

    QAxObject* workBooks = excelApp->querySubObject("WorkBooks");
    workBooks->dynamicCall("Open (const QString&)", QString("C:\\Users\\USER\\Documents\\2ProjectExFile.xlsx"));
}




void MainWindow::on_actionCreating_new_file_after_comparing_triggered()
{
    QXlsx::Document *xlsx_database = new QXlsx::Document("C:\\Users\\USER\\Documents\\2ProjectExFile.xlsx");
    QString cell1,cell2;
    for(int r = 1; r <= 11 ; r++)
    {

           cell1 = xlsx_database->read(r,1).toString();
           cell2=xlsx_database->read(r,2).toString();
           if(cell1!=cell2)
           {

           }
           else
           {
              /*QAxObject * worksheets = workBooks-> querySubObject ( "Sheets");
               QAxObject *sameCells = worksheets->querySubObject("Item(int)", r);
               sameCells->dynamicCall("delete");  permanently deletes the row*/
               xlsx_database->setRowHidden(r,true); //hides


           }
     }

    xlsx_database->saveAs("C:\\Users\\USER\\Documents\\2ProjectExFileAAA.xlsx");
}
