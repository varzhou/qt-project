#ifndef IMPORT_EXCEL_H__
#define IMPORT_EXCEL_H__

#include <QObject>
#include "progressrate.h"

class QAxObject;

class import_excel :public QObject
{
    Q_OBJECT

public:
    import_excel(const QString &filepath, QWidget *parent = 0);
    ~import_excel();

    QList<QStringList> getImportExcelData();

private:
    void readExcel(const QString &filepath);
    int getExcelContentCount(QAxObject *work_book,const int &sheet_count);

private:
    QList<QStringList> m_result;
};


#endif // EXCEL_READ_H
