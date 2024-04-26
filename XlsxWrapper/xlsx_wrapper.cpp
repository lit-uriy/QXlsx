#include "xlsx_wrapper.h"

XlsxWrapper::Book::Book(QString fileName)
    :_valid(false)
{
    _protocolFile = new QXlsx::Document(fileName);
    if (!_protocolFile->isLoadPackage()) {
        qDebug() << "XlsxWrapper::Book::Book() ERROR on open XLSX-file (" << fileName << ")";
        return;
    }
    _valid = true;
}
