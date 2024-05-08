#include "xlsx_wrapper.h"

#include <QDebug>

#if 0
    #define DEBUG qDebug()
#else
    #define DEBUG  QMessageLogger(0, 0, 0).noDebug()
#endif

namespace XlsxWrapper {

//==================================================================================================
//                            Book
//==================================================================================================

Book::Book(QString fileName)
    :_valid(false)
{
    _bookFile = new QXlsx::Document(fileName);
    if (!_bookFile->isLoadPackage()) {
        _errorText = QString("Book::Book() ERROR on open XLSX-file (%1)").arg(fileName);
        DEBUG << _errorText;
        return;
    }
    _valid = true;

    QList<QString> names = _bookFile->sheetNames();
    foreach (QString name, names) {
        QXlsx::Worksheet *ss = dynamic_cast<QXlsx::Worksheet *>(_bookFile->sheet(name));
        if (!ss){
            continue; // QXlsx::Worksheet может не быть, эсли это лист-диаграмма и т.п.
        }
        Sheet *s = new Sheet(name, this, ss);
        _sheets.insert(name, s);
    }
}

Book::~Book()
{
    foreach (Sheet *sh, _sheets) {
        delete sh;
    }
    _sheets.clear();

    delete _bookFile;
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

Cell* Sheet::cell(int rowNumber, int columnNumber)
{
    Cell *c = nullptr;

    // Сначала поищем во внутренней таблице
    if (_cells.contains(rowNumber - 1)){
        QHash<int, Cell*> row = _cells.value(rowNumber - 1);
        if (row.contains(columnNumber - 1)){
            c = row.value(columnNumber - 1);
            return c;
        }
    }

    // Если, не нашли, то ищем в XLSX-е
    QXlsx::Cell *outCell = _xlsxSheet->cellAt(rowNumber, columnNumber);
    if (outCell){
        // Создадим свою ячеку
        c = new Cell();
        c->_rowNumber = rowNumber;
        c->_columnNumber = columnNumber;
        c->_sheet = this;
        c->_data = outCell->value();
        // и поместим её во внутренюю таблицу
        _cells[rowNumber - 1].insert(columnNumber - 1, c);
    }

    return c;
}

Cell* Sheet::findCell(QString text, FindRules fr, Qt::CaseSensitivity cs)
{

    Cell *outCell = nullptr;

    for (int row = 1; row <= _maxRow; ++row) {
        for (int col = 1; col <= _maxCol; ++col) {
            Cell *c = cell(row, col);
            if (!c){
//                DEBUG << "Cell(" << row << "," << col << ") is NULL";
                continue;
            }

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

bool Cell::isSpaned()
{
    QList<QXlsx::CellRange> ranges = _sheet->toXlsx()->mergedCells();
    foreach (QXlsx::CellRange range, ranges) {
        QXlsx::CellReference tl = range.topLeft();
        QXlsx::CellReference br = range.bottomRight();
        bool inBetweenColumn = (tl.column() <= _columnNumber) && (br.column() >= _columnNumber);
        bool inBetweenRow = (tl.row() <= _rowNumber) && (br.row() >= _rowNumber);
        if (inBetweenColumn && inBetweenRow)
            return true;
    }
    return false;
}

} // namespace XlsxWrapper
