#ifndef EXPORT_EXCEL_H
#define EXPORT_EXCEL_H

#include "QObject"
#include "progressrate.h"

class QAxObject;
class ProgressRate;


class export_excel : public QObject
{
    Q_OBJECT

public:
    export_excel(const QList<QStringList> &storeinfo,const QString &storagepath, QWidget *parent = 0);
    ~export_excel();

    enum ExportError
    {
        NoError = 0,
        NewFileError,
        StoreInfoNull,
        TableInfoNotMatch,
        FileExists,
    };

    ExportError exportStatus();

private:
    ExportError   m_status;
    ProgressRate *m_pProgress;
    QAxObject    *m_pApp;
    QAxObject    *m_pWorkbooks;
    QAxObject    *m_pWorkbook;
    QAxObject    *m_pSheet;

private:
    bool newExcel(const QString &storagepath);
    void setCellsInfo(const QList<QStringList> &storeinfo);
    void setCellValue(const int &column, const int &row, const QString &value);
    void saveExcel(const QString &filename);
};

#endif // EXCEL_WRITE_H
