#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include "excelengine.h"
#include <QAxObject>
#include <QMessageBox>
#include <QDebug>
#include <QFileDialog>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_ReadExcel_clicked();

    void on_Open_clicked();

    void on_btn_in_clicked();

    void on_btn_out_clicked();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
