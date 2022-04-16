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

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.doc.table.OdfTableRow;

import fr.paris.lutece.plugins.forms.business.Form;
import fr.paris.lutece.plugins.forms.business.FormHome;
import fr.paris.lutece.plugins.forms.business.FormQuestionResponse;
import fr.paris.lutece.plugins.forms.business.FormResponse;
import fr.paris.lutece.plugins.forms.business.FormResponseHome;
import fr.paris.lutece.plugins.forms.business.FormResponseStep;
import fr.paris.lutece.plugins.forms.business.Step;
import fr.paris.lutece.plugins.forms.business.StepHome;
import fr.paris.lutece.plugins.forms.business.form.FormResponseItem;
import fr.paris.lutece.plugins.forms.business.form.FormResponseItemSortConfig;
import fr.paris.lutece.plugins.forms.business.form.column.FormColumnCell;
import fr.paris.lutece.plugins.forms.business.form.column.IFormColumn;
import fr.paris.lutece.plugins.forms.business.form.filter.FormFilter;
import fr.paris.lutece.plugins.forms.business.form.panel.FormPanel;
import fr.paris.lutece.plugins.forms.export.AbstractFileGenerator;
import fr.paris.lutece.plugins.forms.service.EntryServiceManager;
import fr.paris.lutece.plugins.forms.service.MultiviewFormService;
import fr.paris.lutece.plugins.forms.util.FormMultiviewWorkflowStateNameConstants;
import fr.paris.lutece.plugins.forms.web.entrytype.IEntryDataService;
import fr.paris.lutece.portal.service.i18n.I18nService;
import fr.paris.lutece.util.file.FileUtil;

/**
 * OpenDocument spreadsheet file generator
 */
public class OpenDocumentFileGenerator extends AbstractFileGenerator
{
    private static final String MESSAGE_EXPORT_FORM_TITLE = "forms.export.formResponse.form.title";
    private static final String MESSAGE_EXPORT_FORM_STATE = "forms.export.formResponse.form.state";
    private static final String MESSAGE_EXPORT_FORM_DATE_CREATION = "forms.export.formResponse.form.date.creation";
    private static final String MESSAGE_EXPORT_FORM_DATE_UPDATE = "forms.export.formResponse.form.date.update";

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
    public Path generateFile( ) throws IOException
    {
        Path odfFile = Paths.get( TMP_DIR, getFileName( ) );
        try
        {
            writeExportFile( odfFile );
        }
        catch( Exception e )
        {
            if ( e instanceof IOException )
            {
                throw (IOException) e;
            }
            throw new IOException( e );
        }
        return odfFile;
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
    public boolean isZippable( )
    {
        // ODF is already a zipped file
        return false;
    }

    /**
     * Export the form responses
     * 
     * @param file
     *            file path to export to
     * @throws Exception
     *             if an error occurs
     */
    private void writeExportFile( Path file ) throws Exception
    {
        OdfSpreadsheetDocument document = OdfSpreadsheetDocument.newSpreadsheetDocument( );
        document.getTableList( ).stream( ).forEach( t -> t.remove( ) );
        OdfTable table = OdfTable.newTable( document );
        List<FormResponseItem> searchAllListFormResponseItem = MultiviewFormService.getInstance( ).searchAllListFormResponseItem( _formPanel, _listFormColumn,
                _listFormFilter, _sortConfig );
        Map<Integer, String> mapWorkflowState = searchAllListFormResponseItem.stream( )
                .collect( Collectors.toMap( FormResponseItem::getIdFormResponse, fri -> findWorkflowState( fri ) ) );
        List<FormResponse> listFormResponse = searchAllListFormResponseItem.stream( )
                .map( fri -> FormResponseHome.findByPrimaryKeyForIndex( fri.getIdFormResponse( ) ) ).collect( Collectors.toList( ) );

        Map<String, Integer> mapResponseToColumn = writeHeader( table, listFormResponse );
        Map<Integer, Form> formsByIds = new HashMap<>( );

        int rowIndex = 2;

        for ( FormResponse formResponse : listFormResponse )
        {
            OdfTableRow row = table.getRowByIndex( rowIndex++ );
            Form form = formsByIds.computeIfAbsent( formResponse.getFormId( ), id -> FormHome.findByPrimaryKey( id ) );
            row.getCellByIndex( 0 ).setStringValue( form.getTitle( ) );
            setDateValue( row.getCellByIndex( 1 ), formResponse.getCreation( ) );
            setDateValue( row.getCellByIndex( 2 ), formResponse.getUpdate( ) );
            row.getCellByIndex( 3 ).setStringValue( mapWorkflowState.get( formResponse.getId( ) ) );
            for ( FormResponseStep step : formResponse.getSteps( ) )
            {
                for ( FormQuestionResponse questionResponses : step.getQuestions( ) )
                {
                    if ( !questionResponses.getQuestion( ).isResponseExportable( ) )
                    {
                        continue;
                    }

                    IEntryDataService entryDataService = EntryServiceManager.getInstance( )
                            .getEntryDataService( questionResponses.getQuestion( ).getEntry( ).getEntryType( ) );
                    String responseValue = entryDataService.responseToStrings( questionResponses ).stream( ).collect( Collectors.joining( " " ) );
                    int col = mapResponseToColumn
                            .get( questionResponses.getQuestion( ).getId( ) + "_" + questionResponses.getEntryResponse( ).get( 0 ).getIterationNumber( ) );
                    row.getCellByIndex( col ).setStringValue( responseValue );
                }
            }
        }
        document.save( file.toFile( ) );
    }

    /**
     * Sets a date value to a cell
     * 
     * @param cell
     *            the cell
     * @param timestamp
     *            the date
     */
    private void setDateValue( OdfTableCell cell, Timestamp timestamp )
    {
        Calendar calendar = Calendar.getInstance( );
        calendar.setTimeInMillis( timestamp.getTime( ) );
        cell.setDateValue( calendar );
        // setDateValue cuts off the time, so set it explicitly
        cell.getOdfElement( ).setOfficeDateValueAttribute( DateTimeFormatter.ISO_DATE_TIME.format( timestamp.toLocalDateTime( ) ) );
        cell.setDisplayText( new SimpleDateFormat( "yyyy-MM-dd HH:mm:ss" ).format( timestamp ) );
    }

    /**
     * Writes the header and collect columns position information.
     * 
     * Columns are identified by the question id and the iteration number
     * 
     * @param table
     *            the table
     * @param listFormResponse
     *            the responses to export
     * @return a map from question id to column index
     */
    private Map<String, Integer> writeHeader( OdfTable table, List<FormResponse> listFormResponse )
    {
        List<String> columns = new LinkedList<>( );
        Map<String, String> columnTitles = new HashMap<>( );
        Map<Integer, String> lastIdForQuestion = new HashMap<>( );
        Map<Integer, Set<String>> columnsByStep = new HashMap<>( );
        Map<Integer, String> stepTitles = new HashMap<>( );
        List<Integer> steps = new LinkedList<>( );

        columns.add( MESSAGE_EXPORT_FORM_TITLE );
        columnTitles.put( MESSAGE_EXPORT_FORM_TITLE, I18nService.getLocalizedString( MESSAGE_EXPORT_FORM_TITLE, I18nService.getDefaultLocale( ) ) );
        columns.add( MESSAGE_EXPORT_FORM_DATE_CREATION );
        columnTitles.put( MESSAGE_EXPORT_FORM_DATE_CREATION,
                I18nService.getLocalizedString( MESSAGE_EXPORT_FORM_DATE_CREATION, I18nService.getDefaultLocale( ) ) );
        columns.add( MESSAGE_EXPORT_FORM_DATE_UPDATE );
        columnTitles.put( MESSAGE_EXPORT_FORM_DATE_UPDATE, I18nService.getLocalizedString( MESSAGE_EXPORT_FORM_DATE_UPDATE, I18nService.getDefaultLocale( ) ) );
        columns.add( MESSAGE_EXPORT_FORM_STATE );
        columnTitles.put( MESSAGE_EXPORT_FORM_STATE, I18nService.getLocalizedString( MESSAGE_EXPORT_FORM_STATE, I18nService.getDefaultLocale( ) ) );

        final int stepStartingIndex = columns.size( );

        for ( FormResponse formResponse : listFormResponse )
        {
            for ( FormResponseStep step : formResponse.getSteps( ) )
            {
                for ( FormQuestionResponse questionResponses : step.getQuestions( ) )
                {
                    if ( !questionResponses.getQuestion( ).isResponseExportable( ) )
                    {
                        continue;
                    }
                    int questionId = questionResponses.getQuestion( ).getId( );
                    int iterationNumber = questionResponses.getEntryResponse( ).get( 0 ).getIterationNumber( );
                    String columnId = questionId + "_" + iterationNumber;
                    if ( columnTitles.containsKey( columnId ) )
                    {
                        continue;
                    }
                    if ( lastIdForQuestion.containsKey( questionId ) )
                    {
                        int index = columns.indexOf( lastIdForQuestion.get( questionId ) ) + 1;
                        if ( index > columns.size( ) )
                        {
                            columns.add( columnId );
                        }
                        else
                        {
                            columns.add( index, columnId );
                        }
                        columnTitles.put( columnId, questionResponses.getQuestion( ).getTitle( ) + " (" + iterationNumber + ") " );
                    }
                    else
                    {
                        columns.add( columnId );
                        columnTitles.put( columnId, questionResponses.getQuestion( ).getTitle( ) );
                    }
                    lastIdForQuestion.put( questionId, columnId );
                    columnsByStep.computeIfAbsent( step.getStep( ).getId( ), id -> new HashSet<>( ) ).add( columnId );
                }
                if ( !stepTitles.containsKey( step.getStep( ).getId( ) ) )
                {
                    Step theStep = StepHome.findByPrimaryKey( step.getStep( ).getId( ) ); // step.getStep() does not populate title
                    stepTitles.put( theStep.getId( ), theStep.getTitle( ) );
                    steps.add( theStep.getId( ) );
                }
            }
        }
        OdfTableRow row = table.getRowByIndex( 0 );
        int currentIndex = stepStartingIndex;
        ListIterator<Integer> stepIdIterator = steps.listIterator( );
        while ( stepIdIterator.hasNext( ) )
        {
            Integer stepId = stepIdIterator.next( );
            row.getCellByIndex( currentIndex + columnsByStep.get( stepId ).size( ) - 1 ).setStringValue( stepTitles.get( stepId ) );
            table.getCellRangeByPosition( currentIndex, 0, currentIndex + columnsByStep.get( stepId ).size( ) - 1, 0 ).merge( );
            currentIndex += columnsByStep.get( stepId ).size( );
        }
        Map<String, Integer> mapResponseToColumn = new HashMap<>( );
        ListIterator<String> iterator = columns.listIterator( );
        row = table.getRowByIndex( 1 );
        while ( iterator.hasNext( ) )
        {
            int index = iterator.nextIndex( );
            String columnID = iterator.next( );
            mapResponseToColumn.put( columnID, index );
            row.getCellByIndex( index ).setStringValue( columnTitles.get( columnID ) );
        }
        return mapResponseToColumn;
    }

    /**
     * Lookup the workflow state for a response
     * 
     * @param formResponseItem
     *            the response
     * @return the workflow state, or the empty String
     */
    private String findWorkflowState( FormResponseItem formResponseItem )
    {
        for ( FormColumnCell cell : formResponseItem.getFormColumnCellValues( ) )
        {
            if ( cell != null )
            {
                Object objWorkflowStateName = cell.getFormColumnCellValueByName( FormMultiviewWorkflowStateNameConstants.COLUMN_WORKFLOW_STATE_NAME );
                if ( objWorkflowStateName != null )
                {
                    return String.valueOf( objWorkflowStateName );
                }
            }
        }
        return "";
    }
}
