#include "xlsx_wrapper.h"

#include <QDebug>

namespace XlsxWrapper {

//==================================================================================================
//                            Book
//==================================================================================================

Book::Book(QString fileName)
    :_valid(false)
{
    _bookFile = new QXlsx::Document(fileName);
    if (!_bookFile->isLoadPackage()) {
        qDebug() << "Book::Book() ERROR on open XLSX-file (" << fileName << ")";
        return;
    }
    _valid = true;

    QList<QString> names = _bookFile->sheetNames();
    foreach (QString name, names) {
        Sheet *s = new Sheet(name, this, dynamic_cast<QXlsx::Worksheet *>(_bookFile->sheet(name)));
        _sheets.insert(name, s);
    }
}

Book::~Book()
{
    foreach (Sheet *sh, _sheets) {
        delete sh;
    }
    _sheets.clear();
}

QList<Sheet*> Book::sheets()
{
    return _sheets.values();
}

QList<QString> Book::sheetNames()
{
    return _sheets.keys();
}


Sheet* Book::sheet(QString asheetName)
{
    return _sheets.value(asheetName);
}



//==================================================================================================
//                            Sheet
//==================================================================================================

Sheet::Sheet(QString aname, Book *abook, QXlsx::Worksheet *axlsxSheet)
    :_name(aname)
    ,_book(abook)
    ,_xlsxSheet(axlsxSheet)
{
    _cellList = _xlsxSheet->getFullCells( &_maxRow, &_maxCol );
}

Sheet::~Sheet()
{
    for (int row = 0; row < _maxRow; ++row) {
        QHash<int, Cell*> r = _cells.value(row);

        for (int col = 0; col < _maxCol; ++col) {
            Cell *c = r.value(col);
            delete c;
        }
    }

    _cells.clear();
}

Cell* Sheet::cell(int rowIndex, int colIndex)
{
    Cell *c = nullptr;

    // Сначала поищем во внутренней таблице
    if (_cells.contains(rowIndex)){
        QHash<int, Cell*> row = _cells.value(rowIndex);
        if (row.contains(colIndex)){
            c = row.value(colIndex);
            return c;
        }
    }

    // Если, не нашли, то ищем в XLSX-е
    QXlsx::Cell *outCell = _xlsxSheet->cellAt(rowIndex, colIndex);
    if (outCell){
        // Создадим свою ячеку
        c = new Cell();
        c->_row = rowIndex;
        c->_col = colIndex;
        c->_sheet = this;
        c->_data = outCell->value();
        // и поместим её во внутренюю таблицу
        _cells[rowIndex].insert(colIndex, c);
    }

    return c;
}

Cell* Sheet::findCell(QString text, FindRules fr, Qt::CaseSensitivity cs)
{

    Cell *outCell = nullptr;

    for (int row = 0; row < _maxRow; ++row) {
        QHash<int, Cell*> r = _cells.value(row);

        for (int col = 0; col < _maxCol; ++col) {
            Cell *c = r.value(col);

            // value of cell
            QString str = c->data().toString();

            switch (fr) {
            case StartsWith:
                if (str.startsWith(text, cs)){
                    return outCell = c;
                }
                break;
            case Contains:
                if (str.contains(text, cs)){
                    return outCell = c;
                }
                break;
            default: // EndsWith
                if (str.endsWith(text, cs)){
                    return outCell = c;
                }
                break;
            }
        }
    }

    return outCell;
}

} // namespace XlsxWrapper
