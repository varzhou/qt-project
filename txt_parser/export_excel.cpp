#include "import_excel.h"
#include "export_excel.h"
#include <QDir>
#include <QFile>
#include <QCoreApplication>
#include <ActiveQt\QAxWidget>
#include <ActiveQt\QAxObject>


export_excel::export_excel(const QList<QStringList> &storeinfo,const  QString &storagepath, QWidget *parent)
    : QObject(parent)
    , m_status(NoError)
    , m_pProgress(new ProgressRate)
{
    int count = storeinfo.size();
	if(count <= 0)
		return;
	
    m_pProgress->initProgress(count+1, "file create...");
    m_pProgress->showProgress(0);

    if (newExcel(storagepath)) {
        m_pProgress->updateDescription(tr("file write..."));


        setCellsInfo(storeinfo);
        saveExcel(storagepath);
        delete m_pApp;
    } else {
        m_status = NewFileError;
        return;
    }

    m_pProgress->releaseProgress();
}

export_excel::~export_excel()
{
    delete m_pProgress;
}

export_excel::ExportError export_excel::exportStatus()
{
    return m_status;
}

bool export_excel::newExcel(const QString &storagepath)
{
    m_pApp = new QAxObject();

    m_pApp->setControl("Excel.Application");
    m_pApp->dynamicCall("SetVisible(bool)", false);

    m_pApp->setProperty("DisplayAlerts", false);

    m_pWorkbooks = m_pApp->querySubObject("Workbooks");

    QFile file(storagepath);
    if (file.exists()) {
        m_status = ExportError::FileExists;
        return false;
    } else {
        m_pWorkbooks->dynamicCall("Add");
        m_pWorkbook = m_pApp->querySubObject("ActiveWorkBook");
    }

    m_pSheet = m_pWorkbook->querySubObject("Sheets(int)",1);
    return true;
}

void export_excel::setCellsInfo(const QList<QStringList> &storeinfo)
{
    m_pProgress->showProgress(1);

    // create row info
    for (int row=1; row< storeinfo.size()+1; ++row) {
        QStringList rowinfo = storeinfo.at(row-1);

        //setCellValue(0, row, QString::number(row));
        for (int col=0; col<rowinfo.size(); ++col) {
            QString info = rowinfo.at(col);
            setCellValue(col+1,row,info);
        }
        m_pProgress->showProgress(row);
    }
}

/*
**	:write value to excel cell
*/
void export_excel::setCellValue(const int &column, const int &row, const QString &value)
{
    QAxObject *range = m_pSheet->querySubObject("Cells(int,int)", row, column);
    QString savevalue = value;
    if (value.size() >= 15) {
        savevalue.insert(0, '\'');//add next line
    }
    range->setProperty("Value", value);
}

void export_excel::saveExcel(const QString &filename)
{
    m_pWorkbook->dynamicCall("SaveAs(const QString &)",
        QDir::toNativeSeparators(filename));

    m_pApp->dynamicCall("Quit(void)");
}
