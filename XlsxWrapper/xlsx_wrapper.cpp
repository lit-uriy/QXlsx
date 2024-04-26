#include "xlsx_wrapper.h"

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

XlsxWrapper::Sheet::Sheet(QString aname, Book abook)
    :_name(anme)
    ,_book(abook)
{
    _cellList = _xlsxSheet->getFullCells( &_maxRow, &_maxCol );
}

XlsxWrapper::Cell* XlsxWrapper::Sheet::cell(int rowIndex, int colIndex)
{
    Cell *c = nullptr;

    for ( int index = 0; index < _cellList.size(); ++index ) {
        QXlsx::CellLocation location = _cellList.at(index); // cell location
        if ((location.row == rowIndex) && (location.col == colIndex)){
            // Вернём найденную ячеку
            с = new Cell();
            c->_row = location.row;
            c->_col = location.col;
            c->_sheet = this;
            c->_data = location.cell->value();
            return c;
        }
    }

    return c;
}

XlsxWrapper::Cell* XlsxWrapper::Sheet::findCell(QString text, FindRules fr = Contains, Qt::CaseSensitivity cs = Qt::CaseSensitive)
{
    QXlsx::Cell *outCell = nullptr;

    for ( int index = 0; index < _cellList.size(); ++index ) {
        QXlsx::CellLocation location = _cellList.at(index); // cell location

        // value of cell
        QVariant var = location.cell->value();
        QString str = var.toString();
        auto ptr = location.cell;

        switch (fr) {
        case StartsWith:
            if (str.startsWith(text, cs)){
                outCell = ptr.get();
                return outCell;
            }
            break;
        case Contains:
            if (str.contains(text, cs)){
                outCell = ptr.get();
                return outCell;
            }
            break;
        default: // EndsWith
            if (str.endsWith(text, cs)){
                outCell = ptr.get();
                return outCell;
            }
            break;
        }
    }

    return outCell;
}
