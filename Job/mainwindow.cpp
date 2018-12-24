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

void MainWindow::on_pushButton_clicked()
{
    QString fileName = QFileDialog::getOpenFileName(this, QString::fromLocal8Bit("打开文件"), "C:\\Users\\Lenovo\\Desktop", tr("*.*"));
    if(fileName.isNull())
    {
        return;
    }
    openExcel(fileName);
}

void MainWindow::openExcel(QString fileName)
{
    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false);
    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open(const QString &)", fileName);

    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");

    int sheet_count = work_sheets->property("Count").toInt();
    QList<QList<QVariant>> res;
    if (sheet_count > 0)
    {
        QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);

        QVariant var = readAll(work_sheet);

        castVariant2ListListVariant(var, res);
    }

    work_book->dynamicCall("Close(Boolean)", false);
    excel.dynamicCall("Quit(void)");
}

QVariant MainWindow::readAll(QAxObject *sheet)
{
    QVariant var;
    if (sheet != NULL && !sheet->isNull())
    {
        QAxObject *usedRange = sheet->querySubObject("UsedRange");
        if (NULL == usedRange || usedRange->isNull())
        {
            return var;
        }
        var = usedRange->dynamicCall("Value");
        delete usedRange;
    }
    return var;
}

void MainWindow::castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant>> &res)
{
    QVariantList varRows = var.toList();
    if (varRows.isEmpty())
    {
        return;
    }
    const int rowCount = varRows.size();
    QVariantList rowData;
    for (int i = 0; i < rowCount; ++i)
    {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
}