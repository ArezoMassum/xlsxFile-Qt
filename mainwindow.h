#ifndef MAINWINDOW_H
#define MAINWINDOW_H


#include <QMainWindow>
#include "qaxobject.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:

    void on_actionOpen_Excel_File_triggered();

    void on_actionCreating_new_file_after_comparing_triggered();

private:
    Ui::MainWindow *ui;
    QAxObject *worksheet;
    QAxObject* workBooks;

};

#endif // MAINWINDOW_H


