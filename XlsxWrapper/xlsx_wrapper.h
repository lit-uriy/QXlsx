#ifndef XLSX_WRAPPER_H
#define XLSX_WRAPPER_H


#include "xlsxdocument.h"

class XlsxWrapper
{
public:

    enum FindRules{
        StartsWith = 0,
        EndsWith,
        Contains
    };

    class Book
    {
        QXlsx::Workbook *_xlsxBook;
        QXlsx::Document *_protocolFile;
        QVector<QXlsx::CellLocation> _cellList;
        bool _valid;

    public:
        Book(QString fileName);
        bool isValid();
        QList<Sheet*> sheets();
        QList<QString> sheetNames();

        QXlsx::Workbook* toXlsx();
    };

    class Sheet
    {
        QXlsx::Worksheet *xlsxSheet;

    public:
        QXlsx::Worksheet* toXlsx();

        Cell* findCell(QString text, FindRules fr = Contains, Qt::CaseSensitivity cs = Qt::CaseSensitive);
    };

    class Cell
    {
        int _row, _col;
        Sheet *_sheet;

    public:
        Cell(): _row(-1), _col(-1)
        {}

        QVariant data();
        void setData(QVariant adata);

        int row(){return _row;}
        int col(){return _col;}

        Sheet* sheet();
    };
};

#endif // XLSX_WRAPPER_H
