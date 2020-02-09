#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <data_parser.h>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow,excel_parser
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_pBut_loadfile_clicked();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
