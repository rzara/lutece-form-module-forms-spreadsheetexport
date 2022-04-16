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

import java.io.FileOutputStream;
import java.nio.file.Path;
import java.sql.Timestamp;
import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import fr.paris.lutece.plugins.forms.business.form.FormResponseItemSortConfig;
import fr.paris.lutece.plugins.forms.business.form.column.IFormColumn;
import fr.paris.lutece.plugins.forms.business.form.filter.FormFilter;
import fr.paris.lutece.plugins.forms.business.form.panel.FormPanel;
import fr.paris.lutece.util.file.FileUtil;

public class ExcelFileGenerator extends AbstractSpreadsheetFileGenerator
{

    private XSSFWorkbook _workbook;
    private XSSFSheet _sheet;
    private XSSFRow _row;
    private XSSFCellStyle _dateStyle;

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
    protected ExcelFileGenerator( String fileName, FormPanel formPanel, List<IFormColumn> listFormColumn, List<FormFilter> listFormFilter,
            FormResponseItemSortConfig sortConfig, String fileDescription )
    {
        super( FileUtil.normalizeFileName( fileName ), formPanel, listFormColumn, listFormFilter, sortConfig, fileDescription );
    }

    @Override
    public String getFileName( )
    {
        return _fileName + ".xlsx";
    }

    @Override
    public String getMimeType( )
    {
        return ExcelExport.CONSTANT_MIME_TYPE_XLSX;
    }

    @Override
    protected void prepareDocument( ) throws Exception
    {
        _workbook = new XSSFWorkbook( );
        _sheet = _workbook.createSheet( );
        _dateStyle = _workbook.createCellStyle( );
        _dateStyle.setDataFormat( _workbook.createDataFormat( ).getFormat( "yyyy-mm-dd hh:mm" ) );
    }

    @Override
    protected void saveDocument( Path file ) throws Exception
    {
        try ( FileOutputStream out = new FileOutputStream( file.toFile( ) ) )
        {
            _workbook.write( out );
        }
    }

    @Override
    protected void addRow( int rowIndex )
    {
        _row = _sheet.createRow( rowIndex );
    }

    @Override
    protected void setStringValue( int cellIndex, String value )
    {
        _row.createCell( cellIndex ).setCellValue( value );
    }

    @Override
    protected void setDateValue( int cellIndex, Timestamp timestamp )
    {
        _row.createCell( cellIndex ).setCellValue( timestamp.toLocalDateTime( ) );
        _row.getCell( cellIndex ).setCellStyle( _dateStyle );
    }

    @Override
    protected void mergeCells( int colStartIndex, int rowStartIndex, int colEndIndex, int rowEndIndex )
    {
        _sheet.addMergedRegion( new CellRangeAddress( rowStartIndex, rowEndIndex, colStartIndex, colEndIndex ) );
    }

}
