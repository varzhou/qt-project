#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QDebug>
#include "data_parser.h"


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->setWindowTitle(QString("excel parser"));
    this->setMaximumSize(480,340);
    ui->tEdit_show->setTextColor(Qt::red);
}

MainWindow::~MainWindow()
{
    delete ui;
}

/*
 * filenameo:n_pBut_loadfile_clicked
**  desp:open local file
*/
void MainWindow::on_pBut_loadfile_clicked()
{
    QStringList fileNames;
    ui->tEdit_show->clear();
    ui->tEdit_show->setText(QString("please waiting!!!"));

    /*open dialog*/
    QString filename = QFileDialog::getOpenFileName(this, tr("open excel file"), QCoreApplication::applicationDirPath() ,"excel(*.xlsx)");
    if(filename.isEmpty())
        return;

    load_file(filename);
    ui->tEdit_show->setText(QString("open file is ") + filename + QString("file format is ") +
                            QString((true == file_st)?"ok":"bad"));
	export_file(filename);
}

