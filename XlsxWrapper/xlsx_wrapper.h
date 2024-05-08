#ifndef XLSX_WRAPPER_H
#define XLSX_WRAPPER_H


#include "xlsxdocument.h"

namespace XlsxWrapper {


    enum FindRules{
        StartsWith = 0,
        EndsWith,
        Contains
    };

    class Cell;
    class Sheet;

    class Book
    {
        QXlsx::Workbook *_xlsxBook;
        QXlsx::Document *_bookFile;
        bool _valid;
        QString _errorText;
        QMap<QString, Sheet*> _sheets; // <name, sheet pointer>

    public:
        Book(QString fileName);
       ~Book();

        bool isValid() {return _valid;}
        QString errorText() {return _errorText;}

        Sheet* sheet(QString asheetName);
        QList<Sheet*> sheets();
        QList<QString> sheetNames();

        QXlsx::Workbook* toXlsx(){return _xlsxBook;}
    };

    class Sheet
    {
        friend class Book;

        QString _name;
        Book *_book;
        QXlsx::Worksheet *_xlsxSheet;
        QVector<QXlsx::CellLocation> _cellList;
        int _maxRow;
        int _maxCol;
        QHash<int, QHash<int, Cell*> > _cells; // <строка, <столбец, Ячейка> >

        Sheet(QString aname, Book *abook, QXlsx::Worksheet *axlsxSheet);
       ~Sheet();

    public:
        Book* book(){return _book;}

        int maximumRow() {return _maxRow;}
        int maximumCol() {return _maxCol;}

        Cell* cell(int rowNumber, int columnNumber);
//        Cell* cell(QString coordinat); // буквенно цифровые координаты, например: B7, столбец "B", строка 7

        Cell* findCell(QString text, FindRules fr = Contains, Qt::CaseSensitivity cs = Qt::CaseSensitive);

        QXlsx::Worksheet* toXlsx(){return _xlsxSheet;}
    };

    // FIXME: В данный момент это класс только для чтения, т.к. запись в ячейку ничего не сделает.
    class Cell
    {
        friend class Sheet;

        int _rowNumber, _columnNumber;
        Sheet *_sheet;
        QVariant _data;
        Cell(){}

    public:

        QVariant data(){return _data;}
//        void setData(QVariant adata);

        int rowNumber(){return _rowNumber;}
        int columnNumber(){return _columnNumber;}

        bool isSpaned();

        Sheet* sheet() {return _sheet;}
    };


} // namespace XlsxWrapper

#endif // XLSX_WRAPPER_H
