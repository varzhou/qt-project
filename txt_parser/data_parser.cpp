#include "data_parser.h"
#include "import_excel.h"
#include "export_excel.h"
#include <QFile>
#include <QDir>

/*
*  init class
*/
excel_parser::excel_parser()
{
    file_st =false;
}
/*
**  destroy class
*/
excel_parser::~excel_parser(){

}

void excel_parser::export_file(QString outfile)
{
    QString filename = outfile;

    filename.replace(".xlsx","_res.xlsx");
    if(nullptr == filename)
        return;

    qDebug() << filename;

    export_excel excel(export_datainfo, filename);
    export_excel::ExportError error = excel.exportStatus();
    if (error == export_excel::NoError || error == export_excel::TableInfoNotMatch)
        qDebug() << "output file success";
    else if (error ==  export_excel::StoreInfoNull)
        qDebug() << "output file is empty";
    else if (error == export_excel::FileExists)
        qDebug() << "file is exist";
    else
        qDebug() <<"output fail";
}

/*
*   load input excel file
*/
void excel_parser::load_file(QString fileNames){
    if(fileNames.isEmpty()){
        qDebug()<<"file is empty";
        return;
    }
	bool isfirst = false;
	QStringList slist;
	import_excel excel(fileNames);
    QList<QStringList> importData = excel.getImportExcelData();
    for (auto itemsinfo: importData) {//dispose each line excel data
		if(false == isfirst){
			isfirst = true;
			create_new_header(itemsinfo,slist);
			export_datainfo.append(slist);
			continue;
		}
        slist.clear();
        parse_oneline(itemsinfo,slist);
        export_datainfo.append(slist);
	}
    file_st = true;
}


#define EXCEL_MAX_COL   (2048)
#define EXCEL_MAX_SAMEID    (50)

/**/
struct data_id_t{
    unsigned int grp[EXCEL_MAX_COL][EXCEL_MAX_SAMEID];
    bool st[EXCEL_MAX_COL];
    bool flag[EXCEL_MAX_COL];
};

static struct data_id_t export_head;
/*
**	:create new file first row as header
*/
void excel_parser::create_new_header(QStringList & inl, QStringList & outline){
    unsigned int idx = 0,id_cnt = 1,itx = 0;
    bool rflag = false;
    QStringList::iterator it = inl.begin();

    for(;it < inl.end();it++){
        if(idx >= EXCEL_MAX_COL)
            break;
        if(true == ::export_head.flag[idx]){
            idx++;
            continue;
        }

        QStringList::iterator itb = it;
        itb++;
        for(; itb < inl.end(); itb++){
            id_cnt++;
            if(*it == *itb){
                ::export_head.grp[idx][itx] = id_cnt;
                ::export_head.flag[id_cnt] = true;//set a flag as this item is same to previous item
                if(::export_head.st[idx] == false){
                    outline.append(*it);
                    rflag = true;
                }
                ::export_head.st[idx] = true;
                itx++;
            }

        }
        if(false == rflag)
            outline.append(*it);
        rflag = false;
        idx++;id_cnt = 1;itx = 0;
	}
}

/*
**  :dispose data
*/
void excel_parser::parse_oneline(QStringList & in, QStringList & outline){
	unsigned int row_idx = 0,idx = 0;
	bool dflag = false;
	
    for (auto itemsinfo: in) {//dispose each line excel data
    	if(false == ::export_head.st[row_idx]){
			outline.append(itemsinfo);
		}
		else{
            bool ok = true;
            if((false == itemsinfo.isNull()) && (0.0001f < itemsinfo.toDouble(&ok))){
                outline.append(itemsinfo);//merge data
            }
            else{
                for(idx = 0;idx < EXCEL_MAX_SAMEID;idx++){
                    QString cellval = in.at(::export_head.grp[row_idx][idx]);
                    if(false == cellval.isNull()){
                        if(false == dflag){
                            outline.append(itemsinfo);//merge data
                            dflag = true;
                        }
                    }
                    else{//curretn cell is empty
                        outline.append(itemsinfo);//merge data
                    }
                }
            }
		}
		row_idx++;
	}
}

