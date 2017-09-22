#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_ReadExcel_clicked()
{
    QAxObject *excel = NULL;
    QAxObject *workbooks = NULL;
        QAxObject *workbook = NULL;
        excel = new QAxObject("Excel.Application");
        if (!excel)
        {
            QMessageBox::critical(this, "错误信息", "EXCEL对象丢失");
            return;
        }
        excel->dynamicCall("SetVisible(bool)", false);
        workbooks = excel->querySubObject("WorkBooks");
        workbook = workbooks->querySubObject("Open(QString, QVariant)", QString(tr("F:\\fordebugging\\test.xls")));//Filename
        QAxObject * worksheet = workbook->querySubObject("WorkSheets(int)", 1);//打开第一个sheet
     //QAxObject * worksheet = workbook->querySubObject("WorkSheets");//获取sheets的集合指针
     //int intCount = worksheet->property("Count").toInt();//获取sheets的数量

        QAxObject * usedrange = worksheet->querySubObject("UsedRange");//获取该sheet的使用范围对象
        QAxObject * rows = usedrange->querySubObject("Rows");
        QAxObject * columns = usedrange->querySubObject("Columns");
        /*获取行数和列数*/
        int intRowStart = usedrange->property("Row").toInt();
        int intColStart = usedrange->property("Column").toInt();
        int intCols = columns->property("Count").toInt();
        int intRows = rows->property("Count").toInt();
        /*获取excel内容*/
        for (int i = intRowStart; i < intRowStart + intRows; i++)  //行
        {
            for (int j = intColStart; j < intColStart + intCols; j++)  //列
            {
                QAxObject * cell = worksheet->querySubObject("Cells(int,int)", i, j );  //获取单元格
               // qDebug() << i << j << cell->property("Value");         //*****************************出问题!!!!!!
       qDebug() << i << j <<cell->dynamicCall("Value2()").toString(); //正确
            }
        }
     workbook->dynamicCall("Close (Boolean)", false);
     //同样，设置值，也用dynamimcCall("SetValue(const QVariant&)", QVariant(QString("Help!")))这样才成功的。。
     //excel->dynamicCall("Quit (void)");
     delete excel;//一定要记得删除，要不线程中会一直打开excel.exe
}

void MainWindow::on_Open_clicked()
{
    ExcelEngine excel; //创建excl对象
    excel.Open(QObject::tr("F:\\fordebugging\\Test.xls"),1,false); //打开指定的xls文件的指定sheet，且指定是否可见

    int num = 0;
    for (int i=1; i<=10; i++)
    {
        for (int j=1; j<=10; j++)
        {
           excel.SetCellData(i,j,++num); //修改指定单元数据
        }
    }

    QVariant data = excel.GetCellData(1,1); //访问指定单元格数据
    //QString data = excel.GetCellData(1,1).toString();
    data = excel.GetCellData(2,2);
    data = excel.GetCellData(3,3);

    excel.SetCellData(4,4,111);//QVariant(tr("111"))
    excel.SetCellData(5,4,"222");


    excel.Save(); //保存
    excel.Close();
}

void MainWindow::on_btn_in_clicked()//导入数据到tablewidget中
{
    //getOpenFileName
    QString filename = QFileDialog::getOpenFileName(this,
    tr("Save Excel"),
    "",
    tr("*.xls;; *.xlt;; *.xlsx;; *.xlsm;; *.xltx")); //选择路径
    if(filename.isEmpty())
    {
        return;
    }


    ExcelEngine excel(filename);//QObject::tr("F:\\fordebugging\\Import.xls")
    excel.Open();
    excel.ReadDataToTable(ui->tableWidget); //导入到widget中
    excel.Close();
}

void MainWindow::on_btn_out_clicked()//把tablewidget中的数据导出到excel中
{
    QString filename = QFileDialog::getSaveFileName(this,
    tr("Save Excel"),
    "",
    tr("*.xls;; *.xlt;; *.xlsx;; *.xlsm;; *.xltx")); //选择路径
    if(filename.isEmpty())
    {
        return;
    }

    ExcelEngine excel(filename);//QObject::tr("F:\\fordebugging\\Export.xls")
    excel.Open();
    excel.SaveDataFrTable(ui->tableWidget); //导出报表
    excel.Close();
}
