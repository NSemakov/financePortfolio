classdef ExcelReport < handle
%EXCELREPORT 
% Utiltiy class to create a Excel documentation
%
%   Copyright 2009-2011 The MathWorks, Inc
    
    properties (Access = protected)
        xlsHandle = [];     % Excel automation handle
        enabled = true;     % Reporting enabled/disabled
    end
    
    
    methods (Access = public)

        function this = ExcelReport(filename,sheetname,disable_flag,visible_flag)
        % Create Excel Report object
        %
        % Input arguments:
        % filename:     Filename of Excel report
        % sheetname:    Name of worksheet
        % disable_flag: If this flag is set to 1, reporting is disabled (optional)
        %               May be used to temporarely disable reporting
        % visible_flag: If this flag is set to 1, Excel sheet is visible
        %               during reporting process
        %
        % Last change:   4. January 2010
        % Author:        patric.schenk@mathworks.ch

            % Argument checking
            % -----------------
            if ~exist('filename','var')
                error('ExcelReport:ExcelReport','Please provide a filename');
            end
            if ~ischar(filename)
                error('ExcelReport:ExcelReport','Filename not valid');
            end
            if ~strcmp(filename(end-4:end),'.xlsx') && ~strcmp(filename(end-3:end),'.xls')
                filename = [filename,'.xlsx'];
            end
            if ~exist('sheetname','var')
                error('ExcelReport:ExcelReport','Please provide sheet name');
            end
            if ~ischar(sheetname)
                error('ExcelReport:ExcelReport','Sheet name not valid');
            end
            this.enabled = true;
            if exist('disable_flag','var')
                if disable_flag
                    this.enabled = false;
                end
            end
            if ~exist('visible_flag','var')
                visible_flag = false;
            end

            % Connect to Excel and open/create sheet
            % ---------------------------------------
            try
                if this.enabled
                    % Connect to Excel
                    this.xlsHandle  = actxserver('Excel.Application');
                    this.xlsHandle.visible = visible_flag;
                    % Open file if existing or create new
                    if ~any(filename=='\')    
                        filename = fullfile(pwd,filename);
                    end
                    if exist(filename,'file')
                        % Open file
                        this.xlsHandle.Workbooks.Open(filename);  
                    else
                        % Add new workbook and save file
                        this.xlsHandle.Workbooks.Add;
                        this.xlsHandle.ActiveWorkbook.SaveAs(filename);  
                    end
                    % Activate/add selected sheet
                    found = false;
                    for sheetcount = 1:this.xlsHandle.ActiveWorkBook.Sheets.Count
                        % try to find selected sheet
                        if strcmp(this.xlsHandle.ActiveWorkBook.Sheets.Item(sheetcount).Name,sheetname)
                            this.xlsHandle.ActiveWorkbook.Sheets.Item(sheetcount).Activate;
                            found = true;
                            break;
                        end
                    end
                    if ~found
                        sheet = this.xlsHandle.ActiveWorkBook.Sheets.Add;
                        sheet.Name = sheetname;
                    end
                    % set narrow margins
                    this.xlsHandle.ActiveSheet.PageSetup.HeaderMargin = 0;
                    this.xlsHandle.ActiveSheet.PageSetup.FooterMargin = 0;
                    this.xlsHandle.ActiveSheet.PageSetup.TopMargin = 30;
                    this.xlsHandle.ActiveSheet.PageSetup.BottomMargin = 20;
                    this.xlsHandle.ActiveSheet.PageSetup.LeftMargin = 30;
                    this.xlsHandle.ActiveSheet.PageSetup.RightMargin = 20;
                end    % if enabled
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.xlsHandle.Quit;
                    catch
                    end
                end
                error('ExcelReport:ExcelReport',ME.message);
            end    
        end
        
        function closeReport(this)
        % Close Excel Report

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:CloseReport','Excel handle not valid');
                return
            end

            % Finalize report
            try
                % Save and close Excel report
                this.xlsHandle.ActiveSheet.Range('A1').Select;
                this.xlsHandle.ActiveWorkbook.Save; 
                this.xlsHandle.ActiveWorkbook.Close(false);
                this.xlsHandle.Quit;
                this.xlsHandle = [];  % this removes Excel task

            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.xlsHandle.Quit;
                        this.xlsHandle = [];
                    catch
                    end
                end
                error('ExcelReport:CloseReport',ME.message);
            end        
        end
        
        function insertText(this,range,text,options)
        % Write a text to Excel Report
        %
        % Input arguments:
        % range:   Range to write text (can be multiple cells)
        %          - e.g. 'A3:C3','B4',[4,3], [5,5,32,6] (equals 'E5:F32')
        % text:    Text to add to cell, one of:
        %          - String for single line
        %          - Cell vector containing strings for multiple lines
        % options: Options structure (optional)
        %          - Call getDefaultTextOptions() to retrieve structure setup

            % Get enabled flag
            if this.enabled == false
                return
            end
            % check arguments
            if nargin < 3
                warning('ExcelReport:InsertText','Not enough input arguments, no text inserted');
                return
            end
            if ~ischar(range) && ~isnumeric(range)
                warning('ExcelReport:InsertText','Range not valid, no text inserted');
                return
            end
            if isnumeric(range)
                % replace row/column indeces by corresponding range
                switch numel(range)
                    case 2
                        range = this.index2range(range(1),range(2));
                    case 4
                        range = [this.index2range(range(1),range(2)),':',this.index2range(range(3),range(4))];
                    otherwise
                        warning('ExcelReport:InsertText','Numeric range not valid, no text inserted');
                        return
                end
            end
            if ~ischar(text) && ~iscell(text)
                warning('ExcelReport:InsertText','Text format not valid, no text inserted');
                return
            end
            if ~exist('options','var')
                options = this.getDefaultTextOptions;
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:InsertText','Excel handle not valid');
                return
            end

            % Write data to Excel
            % -------------------
            try 
                % If text is a cell vector, create single string with linefeeds
                if iscell(text)
                    tmp = text{1};
                    for i = 2:length(text)
                        tmp = [tmp,char(10),text{i}];
                    end
                    text = tmp;
                end
                % Write text
                r = this.xlsHandle.ActiveSheet.Range(range);
                r.Select;
                r.Merge;
                r.Value = text;
                if ~isempty(options)
                    r.Font.Name = options.FontName;
                    r.Font.Size = options.FontSize;
                    r.Font.Bold = options.Bold;
                    r.Font.Color = sum(round(options.FGColor*255).*[256*256,256,1]);
                    r.Interior.Color = sum(round(options.BGColor*255).*[256*256,256,1]);    
                    r.Borders.LineStyle = options.Grid;
                    if options.Grid
                        r.Borders.Weight = 1;
                        r.Borders.Color = sum(round(options.GridColor*255).*[256*256,256,1]);
                    end
                    switch options.HorizontalAlignment
                        case 'left'
                            r.HorizontalAlignment = 2;
                        case 'center'
                            r.HorizontalAlignment = 3;
                        case 'right'
                            r.HorizontalAlignment = 4;
                        otherwise
                            error('ExcelReport:InsertText','Horizontal must be one of left/center/right');
                    end
                    switch options.VerticalAlignment
                        case 'top'
                            r.VerticalAlignment = 1;
                        case 'middle'
                            r.VerticalAlignment = 2;
                        case 'bottom'
                            r.VerticalAlignment = 3;
                        otherwise
                            error('ExcelReport:InsertText','Horizontal must be one of top/middle/bottom');
                    end
                end
                % autofit row height
                r.Select;
                this.xlsHandle.Selection.EntireRow.AutoFit;
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:InsertText',ME.message);
            end
        end        
        
        function insertTable(this,range,data,title,options)
        % Write a table to Excel Report
        %
        % Input arguments:
        % range:   Upper left range of table in worksheet 
        %          - e.g. 'A3','B4',[4,3] (equals 'C4')
        % data:    Data to be written
        % title:   Titel of table (optional)
        % options: Options structure (optional)
        %          - call getDefaultTableOptions() to retreive structure setup

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if nargin < 3
                warning('ExcelReport:InsertTable','Not enough input arguments, no table inserted');
                return
            end
            if ~ischar(range) && ~isnumeric(range)
                warning('ExcelReport:InsertTable','Range not valid, no table inserted');
                return
            end
            if isnumeric(range)
                % replace row/column indeces by corresponding range
                switch numel(range)
                    case 2
                        range = this.index2range(range(1),range(2));
                    otherwise
                        warning('ExcelReport:InsertTable','Numeric range not valid, no text inserted');
                        return
                end
            end
            if ~iscell(data) && ~isnumeric(data)
                warning('ExcelReport:InsertTable','Data not valid, no table inserted');
                return
            end
            if ~exist('title','var')
                title = [];
            elseif ~ischar(title)
                title = [];
            end
            if ~exist('options','var')
                options = this.getDefaultTableOptions;
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:InsertTable','Excel handle not valid');
                return
            end

            % Write data to Excel
            % -------------------
            try
                % Get row/col index of range to start with
                [row0,col0] = this.range2index(range);
                % Get size of table (without title)
                [rows,cols] = size(data);
                % do we need to merge cells?
                if ~isempty(options) && ~isempty(options.MergeNumberOfColumns)
                    if numel(options.MergeNumberOfColumns) == cols
                        merge_cols = options.MergeNumberOfColumns;
                    else
                        warning('ExcelReport:InsertTable','Number of elements in "MergeNumberOfColumns" must match number of columns in table data');
                        return
                    end
                else
                    merge_cols = ones(1,cols);
                end
                % Write title (using insertText)
                if ~isempty(title)
                    curr_range = [range,':',this.index2range(row0,col0+sum(merge_cols)-1)];
                    this.insertText(curr_range,title,options.Title);
                    % move start range one row down
                    row0 = row0+1;
                end
                % Write table content
                if ~iscell(data)
                    data = num2cell(data);
                end
                for row = 1:rows
                    for col = 1:cols
                        if merge_cols(col) == 1
                            curr_range =  this.index2range(row0+row-1,col0+sum(merge_cols(1:col))-1);  
                        else
                            curr_range =  [this.index2range(row0+row-1,col0+sum(merge_cols(1:col-1))),':',this.index2range(row0+row-1,col0+sum(merge_cols(1:col))-1)];  
                        end
                        r = this.xlsHandle.ActiveSheet.Range(curr_range);
                        r.Merge;
                        r.Value = data{row,col};
                    end
                end
                curr_range = [this.index2range(row0,col0),':',this.index2range(row0+rows-1,col0+sum(merge_cols)-1)];
                r = this.xlsHandle.ActiveSheet.Range(curr_range);
                if ~isempty(options)
                    r.Font.Name = options.FontName;
                    r.Font.Size = options.FontSize;
                    r.Font.Bold = options.Bold;
                    r.Font.Color = sum(options.FGColor.*[16711680,65280,255]);
                    r.Interior.Color = sum(options.BGColor.*[16711680,65280,255]);
                    if options.MajorGrid
                        r.Borders.Item('xlEdgeTop').LineStyle =1;
                        r.Borders.Item('xlEdgeBottom').LineStyle =1;
                        r.Borders.Item('xlEdgeLeft').LineStyle =1;
                        r.Borders.Item('xlEdgeRight').LineStyle =1;
                        if options.GridColor
                            r.Borders.Item('xlEdgeTop').Color = sum(options.GridColor.*[16711680,65280,255]);
                            r.Borders.Item('xlEdgeBottom').Color = sum(options.GridColor.*[16711680,65280,255]);
                            r.Borders.Item('xlEdgeLeft').Color = sum(options.GridColor.*[16711680,65280,255]);
                            r.Borders.Item('xlEdgeRight').Color = sum(options.GridColor.*[16711680,65280,255]);
                            r.Borders.Item('xlEdgeTop').Weight = 1;
                            r.Borders.Item('xlEdgeBottom').Weight = 1;
                            r.Borders.Item('xlEdgeLeft').Weight = 1;
                            r.Borders.Item('xlEdgeRight').Weight = 1;
                        end
                    end
                    if options.MinorGrid
                        % only set these options if there is multiple rows/columns
                        if rows > 1
                            r.Borders.Item('xlInsideHorizontal').LineStyle =1;
                            r.Borders.Item('xlInsideHorizontal').Weight = 1;
                            if options.GridColor
                                r.Borders.Item('xlInsideHorizontal').Color = sum(options.GridColor.*[16711680,65280,255]);
                            end
                        end
                        if cols > 1
                            r.Borders.Item('xlInsideVertical').LineStyle =1;
                            r.Borders.Item('xlInsideVertical').Weight =1;
                            if options.GridColor
                                r.Borders.Item('xlInsideVertical').Color = sum(options.GridColor.*[16711680,65280,255]);
                            end
                        end
                    end
                    switch options.HorizontalAlignment
                        case 'left'
                            r.HorizontalAlignment = 2;
                        case 'center'
                            r.HorizontalAlignment = 3;
                        case 'right'
                            r.HorizontalAlignment = 4;
                        otherwise
                            error('ExcelReport:InsertTable','Horizontal must be one of left/center/right');
                    end
                    switch options.VerticalAlignment
                        case 'top'
                            r.VerticalAlignment = 1;
                        case 'middle'
                            r.VerticalAlignment = 2;
                        case 'bottom'
                            r.VerticalAlignment = 3;
                        otherwise
                            error('ExcelReport:InsertTable','Horizontal must be one of top/middle/bottom');
                    end
                end
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:InsertTable',ME.message);
            end

        end  
        
        function insertFigure(this,range,figure_handle,dimensions)
        % Insert a MATLAB figure or axes into Excel Report
        %
        % Input arguments:
        % range:         Range in worksheet to place the picture
        %                - e.g. 'A3','B4',[4,3] (equals 'C4')
        % figure_handle: Figure handle or axes handle
        %                Note: for axes, if you want to copy multiple
        %                objects just pass a vector of axes handles. 
        %                (first has to be the main axes object)
        % dimensions:    Two element vector with custom width and height (optional)

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if nargin < 3
                warning('ExcelReport:InsertFigure','Not enough input arguments, no figure inserted');
                return
            end
            if ~ischar(range) && ~isnumeric(range)
                warning('ExcelReport:InsertFigure','Range not valid, no figure inserted');
                return
            end
            if isnumeric(range)
                % replace row/column indeces by corresponding range
                switch numel(range)
                    case 2
                        range = this.index2range(range(1),range(2));
                    otherwise
                        warning('ExcelReport:InsertFigure','Numeric range not valid');
                        return
                end
            end
            if ~any(ishghandle(figure_handle))
                warning('ExcelReport:InsertFigure','Figure handle not valid, no figure inserted');
                return
            end
            if exist('dimensions','var')
                if ~isnumeric(dimensions) || length(dimensions) ~= 2
                    warning('ExcelReport:InsertFigure','Invalid dimensions property, no figure inserted');
                    return
                end
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:InsertFigure','Excel handle not valid');
                return
            end

            % Write figure to Excel
            % ----------------------
            try
                % do not close figure
                close_me = false;
                % use pixel units
                set(figure_handle,'Units','pixels');
                % if handle points to a axes object, first copy content
                % into a new invisible figure
                if ishghandle(figure_handle(1),'axes')
                    % remember to close extra figure at the end
                    close_me = true;
                    % get content and position of original plot
                    axes_position = get(figure_handle(1),'Position');
                    axes_tightinset = get(figure_handle(1),'TightInset');
                    % create new figure
                    f = figure('visible','off');
                    set(f,'Position',[100, ...  %size of figure + buffer of 15pixels
                                      100, ...
                                      axes_position(3)+axes_tightinset(1)+axes_tightinset(3)+15, ...
                                      axes_position(4)+axes_tightinset(2)+axes_tightinset(4)+15]);
                    % copy axes object(s)
                    ax = [];
                    for i = 1:length(figure_handle)
                        ax(i) = copyobj(figure_handle(i),f);
                    end
                    % position main axes object
                    new_axes_position = get(ax(1),'Position');
                    new_axes_tightinset = get(ax(1),'TightInset');
                    set(ax(1),'Position',[new_axes_tightinset(1:2)+5,new_axes_position(3:4)]); % buffer of 5 pixels
                    %set(ax(1),'Units','normalized');
                    % position other objects
                    % compute offset between old and new position of main object
                    new_axes_position = get(ax(1),'Position');
                    offset = new_axes_position(1:2) - axes_position(1:2);
                    for i = 2:length(ax)
                        % add offset to all elements
                        pos = get(figure_handle(i),'Position');
                        set(ax(i),'Position',[pos(1:2)+offset,pos(3:4)]);
                    end
                    for i = 1:length(ax)
                        set(ax(i),'Units','Normalized');  % axes will stretch when resizing figure
                    end
                    % use this figure
                    figure_handle = f;
                end
                
                % Resize figure 
                if exist('dimensions','var')
                    pos = get(figure_handle,'Position');
                    set(figure_handle,'Position',[pos(1:2),dimensions(1),dimensions(2)]);
                end
                % Copy content of figure to clipboard
                set(figure_handle,'PaperPositionMode','auto');
                print(figure_handle,'-dmeta','-painters','-r600');  % vector format and 600dpi
                % Restore dimension of figure
                if exist('dimensions','var')
                    set(figure_handle,'Position',pos);
                end
                % Close figure?
                if close_me
                    close(figure_handle);
                end
                % Paste image in clipboard to Excel range
                r = this.xlsHandle.ActiveSheet.Range(range);
                r.Select;
                this.xlsHandle.ActiveSheet.Paste;
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:InsertFigure',ME.message);
            end
        end

        function insertPicture(this,range,filename,dimensions)
        % Insert a picture into Excel Report
        %
        % Input arguments:
        % range:        Range in worksheet to place the picture
        %                - e.g. 'A3','B4',[4,3] (equals 'C4')
        % filename:     Name of image file (e.g. jpg, gif, etc..)
        % dimensions:    Two element vector with custom width and height

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if nargin < 3
                warning('ExcelReport:InsertPicture','Not enough input arguments, no picture inserted');
                return
            end
            if ~ischar(range) && ~isnumeric(range)
                warning('ExcelReport:InsertPicture','Range not valid, no picture inserted');
                return
            end
            if isnumeric(range)
                % replace row/column indeces by corresponding range
                switch numel(range)
                    case 2
                        range = this.index2range(range(1),range(2));
                    otherwise
                        warning('ExcelReport:InsertPicture','Numeric range not valid, no text inserted');
                        return
                end
            end
            if ~ischar(filename)
                warning('ExcelReport:InsertPicture','Filename not valid, no picture inserted');
                return
            end
            if ~exist(filename,'file')
                warning('ExcelReport:InsertPicture','Filename not valid, no picture inserted');
                return
            end
            % make sure we have absolute path to image
            filename = which(filename);
            if ~exist('dimensions','var')
                warning('ExcelReport:InsertPicture','Dimension of image not defined');
                return
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:InsertPicture','Excel handle not valid');
                return
            end

            % Write picture to Excel
            % ----------------------
            try
                % add rectangle shape
                s = this.xlsHandle.ActiveSheet.Shapes.AddShape(1,0,0,dimensions(1),dimensions(2));
                s.Line.Visible = 0;
                % add picture as background
                s.Fill.UserPicture(filename);
                % move to selected range
                s.Left = this.xlsHandle.ActiveSheet.Range(range).Left;
                s.Top = this.xlsHandle.ActiveSheet.Range(range).Top;

            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:InsertPicture',ME.message);
            end
        end
        
        function createPDF(this,filename)
        % Create a pdf file from current Excel report
        %
        % Input arguments:
        % filename:     Name of pdf report

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if nargin < 2
                warning('ExcelReport:CreatePDF','Please enter filename');
                return
            end
            if ~ischar(filename)
                warning('ExcelReport:CreatePDF','Filename not valid');
                return
            end
            % add current directory if filename is relative
            if isempty(fileparts(filename))
                filename = fullfile(pwd,filename);
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:CreatePDF','Excel handle not valid');
                return
            end

            % Create pdf
            % ----------
            try
                % Use Excel function 'ExportAsFixedFormat' to convert selection
                xlTypePDF = 0;
                xlQualityStandard = 0;
                IncludeDocProperties = true;
                IgnorePrintAreas = false;
                this.xlsHandle.ActiveSheet.ExportAsFixedFormat(xlTypePDF,filename,xlQualityStandard,IncludeDocProperties,IgnorePrintAreas);
                
                % Note: if shapes are not rescaled in the pdf, install
                % hotfix KB973402. It's a MS bug.
               
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:CreatePDF',ME.message);
            end
        end
        
        function setColumnWidth(this,range,width)
        % Set the width of one or multiple columns
        %
        % Input arguments:
        % range: Range in worksheet (e.g. E:F or A)
        % width: New width of columns in range (leave empty for autofit)

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if nargin < 2
                warning('ExcelReport:SetColumnWidth','Not enough input arguments, column width not set');
                return
            end
            if ~ischar(range)
                warning('ExcelReport:SetColumnWidth','Range not valid, column width not set');
                return
            end
            if ~exist('width','var')
                width = [];  % autofit
            end
            if ~isnumeric(width)
                warning('ExcelReport:SetColumnWidth','Width not valid, column width not set');
                return
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:SetColumnWidth','Excel handle not valid');
                return
            end

            % Set width
            % ---------
            try
                % we need a little trick here 
                % select a cell that is (most probably :) not merged with other cells 
                % (row 10000). Otherwise unwanted columns might be changed.

                % build range to select, either a single column or multiple
                [r1,rem_str] = strtok(range,':');
                range = [r1,'10000'];
                if ~isempty(rem_str)
                    r2 = rem_str(2:end);
                    range = [range,':',r2,'10000'];
                end
                this.xlsHandle.ActiveSheet.Range(range).Select;
                if isempty(width)
                    this.xlsHandle.Selection.EntireColumn.AutoFit;
                else
                    this.xlsHandle.Selection.columnWidth = width;
                end
                this.xlsHandle.ActiveSheet.Range('A1').Select;
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:SetColumnWidth',ME.message);
            end
        end

        function setRowHeight(this,range,height)
        % Set the height of one or multiple rows
        %
        % Input arguments:
        % range: Range in worksheet (e.g. 34:50 or 23)
        % height: New height of rows in range (leave empty for autofit)

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if nargin < 2
                warning('ExcelReport:setRowHeight','Not enough input arguments, row height not set');
                return
            end
            if ~ischar(range)
                warning('ExcelReport:setRowHeight','Range not valid (must be string type), row height not set');
                return
            end
            if ~exist('height','var')
                height = [];  % autofit
            end
            if ~isnumeric(height)
                warning('ExcelReport:setRowHeight','Height not valid, row height not set');
                return
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:setRowHeight','Excel handle not valid');
                return
            end

            % Set height
            % ----------
            try
                % we need a little trick here 
                % select a cell that is (most probably :) not merged with other cells 
                % (column ZZ). Otherwise unwanted rows might be changed.

                % build range to select, either a single column or multiple
                [c1,rem_str] = strtok(range,':');
                range = ['ZZ',c1];
                if ~isempty(rem_str)
                    c2 = rem_str(2:end);
                    range = [range,':','ZZ',c2];
                end
                this.xlsHandle.ActiveSheet.Range(range).Select;
                if isempty(height)
                    this.xlsHandle.Selection.EntireRow.AutoFit;
                else
                    this.xlsHandle.Selection.rowHeight = height;
                end
                this.xlsHandle.ActiveSheet.Range('A1').Select;
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:setRowHeight',ME.message);
            end
        end        
        
        function setOrientation(this,orientation)
        % Set paper orientation
        %
        % Input arguments:
        % orientation: One of 'Portrait' / 'Landscape' or 0 / 1 respectively

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if nargin < 2
                warning('ExcelReport:SetOrientation','Not enough input arguments, orientation not set');
                return
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:SetOrientation','Excel handle not valid');
                return
            end

            % Set orientation
            try
                if isnumeric(orientation)
                    if orientation == 0
                        orientation = 'xlPortrait';
                    else
                        orientation = 'xlLandscape';
                    end
                else
                    if strcmp(orientation,'Portrait')
                        orientation = 'xlPortrait';
                    else
                        orientation = 'xlLandscape';
                    end
                end
                this.xlsHandle.Activesheet.PageSetup.Orientation = orientation;
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:SetOrientation',ME.message);
            end

        end
        
        function setFitToPage(this,fittopageswide,fittopagestall)
        % Setup worksheet to fit to a number of pages
        %
        % Input arguments:
        % fittopageswide: Number of pages wide (default = 1)
        % fittopagestall: Number of pages tall (default = 1)

            % Get enabled flag
            if this.enabled == false
                return
            end
            % Get enabled flag
            if this.enabled == false
                return
            end
            % Argument checking
            if ~exist('fittopageswide','var')
                fittopageswide = 1; 
            end
            if ~isnumeric(fittopageswide)
                warning('ExcelReport:setFitToOnePage','Number of pages width not valid');
                return
            end
            if ~exist('fittopagestall','var')
                fittopagestall = 1; 
            end
            if ~isnumeric(fittopagestall)
                warning('ExcelReport:setFitToOnePage','Number of pages height not valid');
                return
            end
            % Get handle to Excel
            if ~isa(this.xlsHandle,'COM.Excel_Application')
                warning('ExcelReport:setFitToOnePage','Excel handle not valid');
                return
            end

            try
                % Change page setup to fit-to-one-page 
                this.xlsHandle.ActiveSheet.PageSetup.Zoom = false; 
                this.xlsHandle.ActiveSheet.PageSetup.FitToPagesWide = fittopageswide;
                this.xlsHandle.ActiveSheet.PageSetup.FitToPagesTall = fittopagestall;
                
            catch ME
                % connection to Excel failed
                % try to close Excel
                if isa(this.xlsHandle,'COM.Excel_Application')
                    try
                        this.closeReport();
                    catch
                    end
                end
                error('ExcelReport:setFitToOnePage',ME.message);
            end

        end
         
     end
    
    
    methods (Access = public, Static)
        
        function text_options = getDefaultTextOptions
        % --- Create default table options structure ---
            text_options.Grid = false;
            text_options.GridColor = [0,0,0];
            text_options.BGColor = [1,1,1];
            text_options.FGColor = [0,0,0];
            text_options.FontSize = 10;
            text_options.FontName = 'Arial';
            text_options.Bold = false;
            text_options.HorizontalAlignment = 'left';
            text_options.VerticalAlignment = 'top';     
        end
        
        function table_options = getDefaultTableOptions
        % Create default table options structure
            table_options.Title.Grid = false;
            table_options.Title.GridColor = [0,0,0];
            table_options.Title.BGColor = [1,1,1];
            table_options.Title.FGColor = [0,0,0];
            table_options.Title.FontSize = 10;
            table_options.Title.FontName = 'Arial';
            table_options.Title.Bold = false;
            table_options.Title.HorizontalAlignment = 'center';
            table_options.Title.VerticalAlignment = 'bottom';
            table_options.MajorGrid = false;
            table_options.MinorGrid = false;
            table_options.GridColor = [0,0,0];
            table_options.BGColor = [1,1,1];
            table_options.FGColor = [0,0,0];
            table_options.FontSize = 10;
            table_options.FontName = 'Arial';
            table_options.Bold = false;
            table_options.HorizontalAlignment = 'left';
            table_options.VerticalAlignment = 'bottom';
            table_options.MergeNumberOfColumns = [];
            % Note: if defined, number of elements of MergeNumberOfColumns 
            % matches number of columns in table data. Each element defines
            % the number of columns to merge per row before inserting data.
        end
        
        function range = index2range(row,col)
        % Return range info from row/column index
        %
        % Input arguments:
        % row: Row index
        % col: Column index

            range = '';
            % add column info as letters
            % -> maximum is column ZZ
            if col > 26
                range = [range,char(64+floor((col-1)/26))];
            end
            tmp = mod(col,26);
            if tmp == 0
                tmp = 26;
            end
            range = [range,char(64+tmp)];
            % add row info as number
            range = [range,num2str(row)];
        end
        
        function [row,col] = range2index(range)
        % Extract row and column index from range
        %
        % Input arguments:
        % range:  Range in worksheet (e.g. 'E3')

            index = find(double(range)<65,1);
            row = str2double(range(index:end));
            col_tmp = range(1:index-1);
            % -> maximum is column ZZ
            col = 0;
            if length(col_tmp)>1
                col = col + (double(col_tmp(1))-64)*26;
                col_tmp = col_tmp(2);
            end
            col = col + double(col_tmp)-64;
        end
        
    end
   
end











%         function addSheet(this,sheetname)
%         % Add/Activate selected sheet
%         %
%         % Input arguments:
%         % sheetname: Name of sheet to add
% 
%         % Get enabled flag
%         if this.enabled == false
%             return
%         end
%         % Get handle to Excel
%         if ~isa(this.xlsHandle,'COM.Excel_Application')
%             warning('ExcelReport:AddSheet','Excel handle not valid');
%             return
%         end
%         % Argument checking
%         if nargin < 2
%             warning('ExcelReport:AddSheet','Not enough input arguments, no sheet added');
%             return
%         end
%         if ~ischar(sheetname)
%             warning('ExcelReport:AddSheet','Sheet name not valid, no sheet added');
%             return
%         end
% 
%         % Add/activate sheet
%         % ------------------
%         try 
%             % Try to find sheet
%             found = false;
%             for sheetcount = 1:this.xlsHandle.ActiveWorkBook.Sheets.Count
%                 % try to find selected sheet
%                 if strcmp(this.xlsHandle.ActiveWorkBook.Sheets.Item(sheetcount).Name,sheetname)
%                     this.xlsHandle.ActiveWorkbook.Sheets.Item(sheetcount).Activate;
%                     found = true;
%                     break;
%                 end
%             end
%             if ~found
%                 sheet = this.xlsHandle.ActiveWorkBook.Sheets.Add;
%                 sheet.Name = sheetname;
%             end
%         catch ME
%             % connection to Excel failed
%             % try to close Excel
%             if isa(this.xlsHandle,'COM.Excel_Application')
%                 try
%                     this.closeReport();
%                 catch
%                 end
%             end
%             error('ExcelReport:AddSheet',ME.message);
%         end
%         end       