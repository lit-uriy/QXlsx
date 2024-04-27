#include "xlsx_wrapper.h"

#include <QDebug>

//==================================================================================================
//                            Book
//==================================================================================================

XlsxWrapper::Book::Book(QString fileName)
    :_valid(false)
{
    _bookFile = new QXlsx::Document(fileName);
    if (!_bookFile->isLoadPackage()) {
        qDebug() << "XlsxWrapper::Book::Book() ERROR on open XLSX-file (" << fileName << ")";
        return;
    }
    _valid = true;

    QList<QString> sheetNames();
    foreach (QString name, sheetNames()) {
        Sheet *s = new Sheet(name, this);
        s->_xlsxSheet = dynamic_cast<QXlsx::Worksheet *>(_bookFile->sheet(name));
        _sheets.insert(name, s);
    }
}
QList<XlsxWrapper::Sheet*> XlsxWrapper::Book::sheets()
{
    return _sheets.values();
}

QList<QString> XlsxWrapper::Book::sheetNames()
{
    return _sheets.keys();
}


XlsxWrapper::Sheet* XlsxWrapper::Book::sheet(QString asheetName)
{
    return _sheets.value(asheetName);
}



//==================================================================================================
//                            Sheet
//==================================================================================================

XlsxWrapper::Sheet::Sheet(QString aname, Book *abook)
    :_name(aname)
    ,_book(abook)
{
    _cellList = _xlsxSheet->getFullCells( &_maxRow, &_maxCol );
}

XlsxWrapper::Cell* XlsxWrapper::Sheet::cell(int rowIndex, int colIndex)
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

XlsxWrapper::Cell* XlsxWrapper::Sheet::findCell(QString text, FindRules fr, Qt::CaseSensitivity cs)
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
