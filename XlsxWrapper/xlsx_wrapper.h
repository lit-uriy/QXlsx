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
        QMap<QString, Sheet*> _sheets; // <name, sheet pointer>

    public:
        Book(QString fileName);
       ~Book(); // TODO: сделать удаление книг

        bool isValid() {return _valid;}

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

        Sheet(QString aname, Book *abook);
       ~Sheet(); // TODO: сделать удаление ячеек

    public:
        Book* book(){return _book;}
        Cell* cell(int rowIndex, int colIndex);
//        Cell* cell(QString coordinat); // буквенно цифровые координаты, например: B7, столбец "B", строка 7

        Cell* findCell(QString text, FindRules fr = Contains, Qt::CaseSensitivity cs = Qt::CaseSensitive);

        QXlsx::Worksheet* toXlsx(){return _xlsxSheet;}
    };

    // FIXME: В данный момент это класс только для чтения, т.к. запись в ячейку ничего не сделает.
    class Cell
    {
        friend class Sheet;

        int _row, _col;
        Sheet *_sheet;
        QVariant _data;
        Cell(){}

    public:

        QVariant data(){return _data;}
//        void setData(QVariant adata);

        int row(){return _row;}
        int column(){return _col;}

        Sheet* sheet();
    };


} // namespace XlsxWrapper

#endif // XLSX_WRAPPER_H
