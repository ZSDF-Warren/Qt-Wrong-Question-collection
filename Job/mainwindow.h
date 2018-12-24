#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFileDialog>
#include <ActiveQt/QAxObject>

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
    void on_pushButton_clicked();

private:
    Ui::MainWindow *ui;
    void openExcel(QString fileName);
    QVariant readAll(QAxObject *sheet);
    void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant>> &res);
};

#endif // MAINWINDOW_H
