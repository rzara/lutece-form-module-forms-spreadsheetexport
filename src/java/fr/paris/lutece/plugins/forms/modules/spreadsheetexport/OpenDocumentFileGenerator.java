/*
 * Copyright (c) 2002-2022, City of Paris
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions
 * are met:
 *
 *  1. Redistributions of source code must retain the above copyright notice
 *     and the following disclaimer.
 *
 *  2. Redistributions in binary form must reproduce the above copyright notice
 *     and the following disclaimer in the documentation and/or other materials
 *     provided with the distribution.
 *
 *  3. Neither the name of 'Mairie de Paris' nor 'Lutece' nor the names of its
 *     contributors may be used to endorse or promote products derived from
 *     this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 *
 * License 1.0
 */
package fr.paris.lutece.plugins.forms.modules.spreadsheetexport;

import java.nio.file.Path;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.List;

import org.odftoolkit.odfdom.doc.OdfDocument;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.doc.table.OdfTableRow;

import fr.paris.lutece.plugins.forms.business.form.FormResponseItemSortConfig;
import fr.paris.lutece.plugins.forms.business.form.column.IFormColumn;
import fr.paris.lutece.plugins.forms.business.form.filter.FormFilter;
import fr.paris.lutece.plugins.forms.business.form.panel.FormPanel;
import fr.paris.lutece.util.file.FileUtil;

public class OpenDocumentFileGenerator extends AbstractSpreadsheetFileGenerator
{

    private OdfTable _table;
    private OdfDocument _document;
    private OdfTableRow _row;

    /**
     * Constructor
     * 
     * @param fileName
     *            file name for the export
     * @param formPanel
     *            the form panel
     * @param listFormColumn
     *            the form columns list
     * @param listFormFilter
     *            the form filter
     * @param sortConfig
     *            the sort configuration
     * @param fileDescription
     *            description of the file
     */
    protected OpenDocumentFileGenerator( String fileName, FormPanel formPanel, List<IFormColumn> listFormColumn, List<FormFilter> listFormFilter,
            FormResponseItemSortConfig sortConfig, String fileDescription )
    {
        super( FileUtil.normalizeFileName( fileName ), formPanel, listFormColumn, listFormFilter, sortConfig, fileDescription );
    }

    @Override
    public String getFileName( )
    {
        return _fileName + ".ods";
    }

    @Override
    public String getMimeType( )
    {
        return OpenDocumentExport.CONSTANT_MIME_TYPE_OPENDOCUMENT;
    }

    @Override
    protected void prepareDocument( ) throws Exception
    {
        _document = OdfSpreadsheetDocument.newSpreadsheetDocument( );
        _document.getTableList( ).stream( ).forEach( t -> t.remove( ) );
        _table = OdfTable.newTable( _document );
    }

    @Override
    protected void saveDocument( Path file ) throws Exception
    {
        _document.save( file.toFile( ) );
    }

    @Override
    protected void addRow( int index )
    {
        _row = _table.getRowByIndex( index );
    }

    @Override
    protected void setStringValue( int cellIndex, String value )
    {
        _row.getCellByIndex( cellIndex ).setStringValue( value );
    }

    @Override
    protected void setDateValue( int cellIndex, Timestamp timestamp )
    {
        OdfTableCell cell = _row.getCellByIndex( cellIndex );
        Calendar calendar = Calendar.getInstance( );
        calendar.setTimeInMillis( timestamp.getTime( ) );
        cell.setDateValue( calendar );
        // setDateValue cuts off the time, so set it explicitly
        cell.getOdfElement( ).setOfficeDateValueAttribute( DateTimeFormatter.ISO_DATE_TIME.format( timestamp.toLocalDateTime( ) ) );
        cell.setDisplayText( new SimpleDateFormat( "yyyy-MM-dd HH:mm:ss" ).format( timestamp ) );
    }

    @Override
    protected void mergeCells( int colStartIndex, int rowStartIndex, int colEndIndex, int rowEndIndex )
    {
        _table.getCellRangeByPosition( colStartIndex, rowStartIndex, colEndIndex, rowEndIndex ).merge( );
    }

}
