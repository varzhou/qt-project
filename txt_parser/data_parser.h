#ifndef DATA_PARSER_H
#define DATA_PARSER_H

#include <QByteArrayData>
#include <QDebug>
#include <QAxObject>

class excel_parser
{
public:
    excel_parser();
    ~excel_parser();

public:
    bool file_st;
    void load_file(QString fileNames);
    void export_file(QString outfile);
protected:
    QList<QStringList> export_datainfo;
    void create_new_header(QStringList &in,QStringList &outline);
    void parse_oneline(QStringList &in,QStringList &outline);
};

#endif // TXT_PARSER_H
