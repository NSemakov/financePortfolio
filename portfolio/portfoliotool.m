function varargout = portfoliotool(varargin)
% PORTFOLIOTOOL
%      PORTFOLIOTOOL() 


% Edit the above text to modify the response to help portfoliotool

% Last Modified by GUIDE v2.5 09-Mar-2016 20:31:15

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @portfoliotool_OpeningFcn, ...
                   'gui_OutputFcn',  @portfoliotool_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT
end

% --- Executes just before portfoliotool is made visible.
function portfoliotool_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to portfoliotool (see VARARGIN)

    % Choose default command line output for portfoliotool
    handles.output = hObject;

    % ---------------------------------------------------------------------
    % Main setup
    % ---------------------------------------------------------------------

    % Create portfolio model
    handles.model = Portfolio();
    
    % Selected portfolio 
    % Note: handles.selection is the marker handle. for the portfolio
    % number, use "get(handles.selection,'Userdata');"
    handles.selection  = [];
    
    % Date string format (used to convert string into serial dates)
    handles.datestringformat = 'dd.mm.yyyy';

    % Set renderer
    set(handles.figure,'Renderer','OpenGL')

    % ---------------------------------------------------------------------
    % Tab setup
    % ---------------------------------------------------------------------
    % Create custom tabs
    tab_labels = {'Data Import','Portfolio Optimization','Results'};
    handles.tabcount = length(tab_labels);  % Number of tabs
    for i = 1:handles.tabcount
        eval(['h = handles.axes_tab_',num2str(i),';']);  
        axes(h);
        pos = get(h,'Position');
        set(h,'XLim',[0,pos(3)]);
        set(h,'YLim',[0,pos(4)]);
        set(h,'XTick',[]);
        set(h,'YTick',[]);
        set(h,'XTickLabel',{});
        set(h,'YTickLabel',{});
        patch([0,5,pos(3)-5,pos(3),0],[0,pos(4),pos(4),0,0],[1,1,1]);
        text(pos(3)/2,pos(4)/2+2,tab_labels{i},'HorizontalAlignment','center','FontSize',8,'FontName','MS Sans Serif','Units','pixels');
        c = get(h,'Children');    % make sure we can always click on axes object
        for j = 1:length(c)  
            set(c(j),'HitTest','off');
        end
    end
    
    % ---------------------------------------------------------------------
    % Clear all panels
    % ---------------------------------------------------------------------
    clearImportDataPage(handles);
    clearDataSeriesPage(handles);
    clearPortfolioOptimizationPage(handles);
    clearResultsPage(handles);
    
    % Make data import page current
    tab_handler(1,handles);

    % Update handles structure
    guidata(hObject, handles);

end

% --- Outputs from this function are returned to the command line.
function varargout = portfoliotool_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

end


function edit_dataseries_decayfactor_Callback(hObject, eventdata, handles)
% hObject    handle to edit_dataseries_decayfactor (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

	% Get decay factor
    decayfactor = str2double(get(handles.edit_dataseries_decayfactor,'String'));
    % Handle bad entries
    if isnan(decayfactor) || isempty(decayfactor) || (decayfactor <= 0) || (decayfactor > 100)
        decayfactor = 99;
        set(handles.edit_dataseries_decayfactor,'String','99');
    end
    % Update model
    handles.model.setDecayFactor(decayfactor/100);
    % Update visualization
    updateDataSeriesPage(handles) 
end


% --- Executes during object creation, after setting all properties.
function edit_dataseries_decayfactor_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_dataseries_decayfactor (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


% --- Executes on button press in checkbox_dataseries_decayfactor.
function checkbox_dataseries_decayfactor_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_dataseries_decayfactor (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Get status of checkbox
    state = get(handles.checkbox_dataseries_decayfactor,'Value');

    % Enable/disable decay factor edit field
    if state
        set(handles.edit_dataseries_decayfactor,'Enable','on');
    else
        set(handles.edit_dataseries_decayfactor,'Enable','off');
    end
    
    % If checkbox enabled, trigger update
    % Reset decay factor otherwise
    if state
        % Fire edit field callback
        edit_dataseries_decayfactor_Callback(handles.edit_dataseries_decayfactor,[],handles);
    else
        % Update model
        handles.model.setDecayFactor(1);
        % Update visualization
        updateDataSeriesPage(handles)        
    end
    
end

% --- Clear Import Data Page
function disableImportDataPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % Disable all components on "imported data" page
    set(handles.uitable_importedseries_pricesseries,'Enable','off');
    set(handles.uitable_importedseries_pricesseries,'Data',[]);
    set(handles.uitable_importedseries_pricesseries,'RowName',[]);
    set(handles.uitable_importedseries_pricesseries,'ColumnName',[]);
    set(handles.checkbox_importedseries_usefirstcolasbenchmark,'Enable','off');
    set(handles.checkbox_importedseries_usefirstcolasbenchmark,'Value',0);
    set(handles.button_importedseries_accept,'Enable','off');
    pause(0.05);  % give some time to update
    
end


% --- Clear Import Data Page
function clearImportDataPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % Disabe all import panels
    set(handles.uipanel_dataimport_workspace,'Visible','off');
    set(handles.uipanel_dataimport_datafeed,'Visible','off');
    set(handles.uipanel_dataimport_xlsfile,'Visible','off');
    
    % Import from Workspace panel

    % Datafeed panel
    [yyyy,mm,dd] = datevec(today);  % predefine date range
    set(handles.edit_dataimport_datafeed_startdate,'String',datestr([yyyy-1,mm,dd,0,0,0],handles.datestringformat));
    set(handles.edit_dataimport_datafeed_enddate,'String',datestr([yyyy,mm,dd,0,0,0],handles.datestringformat));
    set(handles.uipanel_dataimport_datafeed_download,'Visible','off');  % hide download panel
    
    % Excel Import panel
    
    % Activate "MATLAB Workspace" panel
    set(handles.popupmenu_dataimport_selectsource,'Value',3);
    popupmenu_dataimport_selectsource_Callback(handles.popupmenu_dataimport_selectsource, [], handles);

    % Disable all components on "imported data" page
    set(handles.uitable_importedseries_pricesseries,'Enable','off');
    set(handles.uitable_importedseries_pricesseries,'Data',[]);
    set(handles.checkbox_importedseries_usefirstcolasbenchmark,'Enable','off');
    set(handles.checkbox_importedseries_usefirstcolasbenchmark,'Value',0);
    set(handles.button_importedseries_accept,'Enable','off');
    
end

% --- Clear Data Series Page
function clearDataSeriesPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % Reset all data series controls
    set(handles.uitable_dataseries_assetselection,'ColumnName',[],'RowName',[],'Data',[]);
    cla(handles.axes_dataseries_returnseries);
    reset(handles.axes_dataseries_returnseries);
    set(handles.axes_dataseries_returnseries,'Visible','off');
    set(handles.checkbox_dataseries_decayfactor,'Value',0);
    set(handles.checkbox_dataseries_logreturns,'Value',0);
    
    % Disable controls
    set(handles.checkbox_dataseries_decayfactor,'Enable','off');
    set(handles.checkbox_dataseries_logreturns,'Enable','off');
    
end

% --- Clear Portfolio Optimization Page
function clearPortfolioOptimizationPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % Optimization constraints tables
    set(handles.uitable_portopt_genericconstraints,'Data',[],'ColumnName',[]);

    % Bounds constraints table
    set(handles.uitable_portopt_boundconstraints,'Data',[]);
    set(handles.uitable_portopt_boundconstraints,'ColumnWidth',{50});
    set(handles.uitable_portopt_boundconstraints,'ColumnFormat',{'numeric'});
    set(handles.uitable_portopt_boundconstraints,'ColumnEditable',true);

    % Disable controls
    set(handles.button_portopt_addconstraint,'Enable','off');
    set(handles.button_portopt_computeefficientfrontier,'Enable','off');
    
end

% --- Clear Results Page
function clearResultsPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % For performance reasons, first check if page is already cleared
    if strcmp(get(handles.button_results_createreport,'Enable'),'off')
        % ok, done
        return
    end
    
    % Clear results page
    cla(handles.axes_results_efficientfrontier);
    legend(handles.axes_results_efficientfrontier,'off')
    set(handles.axes_results_efficientfrontier,'Visible','off');
    cla(handles.axes_results_performance);
    legend(handles.axes_results_performance,'off');
    set(handles.axes_results_performance,'Visible','off');
    cla(handles.axes_results_valueatrisk);
    set(handles.axes_results_valueatrisk,'Visible','off');
    axes(handles.axes_results_allocation);
    h = pie(1,{'Select portfolio from efficient frontier'});
    for i = 1:length(h)
        if strcmp(get(h(i),'Type'),'patch')
            set(h(i),'Visible','off')
        end
    end
    set(handles.uitable_results_weights,'Data',[]);
    set(handles.uitable_results_weights,'ColumnName',[]);
    set(handles.uitable_results_metrics,'Data',[]);
    set(handles.uitable_results_metrics,'ColumnFormat',{'char','char'});
    pos = get(handles.uitable_results_metrics,'Position');
    set(handles.uitable_results_metrics,'ColumnWidth',{135,pos(3)-135-4});
    % No portfolio selected
    handles.selection = [];
    guidata(handles.figure,handles);
    
    % Disable controls
    set(handles.edit_results_confidencelevel,'Enable','off');
    set(handles.edit_results_riskfreerate,'Enable','off');
    set(handles.popupmenu_results_valueatrisk,'Enable','off');
    set(handles.button_results_createreport,'Enable','off');
    
end


% --- Update Data Data Series 
function updateDataSeriesPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % Only run if data available
    if isempty(handles.model.getPrices)
        return
    end

    % Get all needed data
    returns         = handles.model.getReturnSeries;
    dates           = handles.model.getDates;
    assetlabels     = handles.model.getPricesLabels;
    [~,~,~,annualized_ret,annualized_rsk] = handles.model.getStatistics(true);  % annualized stats for all assets
    assetselection  = handles.model.getAssetSelection;
    assetsel_ind    = find(assetselection); % we need linear indices
    
    % Asset selection
    % Setup selection table after new data has been imported
    if isempty(get(handles.uitable_dataseries_assetselection,'Data'))
        set(handles.uitable_dataseries_assetselection,'ColumnEditable',[false,true,false,false]);
        set(handles.uitable_dataseries_assetselection,'ColumnFormat',{'char','logical','char','char'});
        set(handles.uitable_dataseries_assetselection,'RowName',[]);
        % set column headers
        set(handles.uitable_dataseries_assetselection,'ColumnName',{'Asset Name','Select','ANL Ret','ANL Vol'});
        % create table content
        data = num2cell([annualized_ret(:),annualized_rsk(:)]);
        for i = 1:numel(data)
            data{i} = [sprintf('%2.2f',data{i}*100),'%'];
        end
        data = [assetlabels(:),repmat({true},length(assetlabels),1),data];
        set(handles.uitable_dataseries_assetselection,'Data',data);
        pos = get(handles.uitable_dataseries_assetselection,'Position');
        if length(assetlabels) < 11
            set(handles.uitable_dataseries_assetselection,'ColumnWidth',{pos(3)-184,60,60,60});
        else
            set(handles.uitable_dataseries_assetselection,'ColumnWidth',{pos(3)-200,60,60,60});
        end
    end
    
    % Update table with annualized assets statistics
    data = get(handles.uitable_dataseries_assetselection,'Data');
    newstatsdata = num2cell([annualized_ret(:),annualized_rsk(:)]);
    for i = 1:numel(newstatsdata)
        newstatsdata{i} = [sprintf('%2.2f',newstatsdata{i}*100),'%'];
    end
    if sum(any(cellfun(@strcmp,newstatsdata,data(:,3:4))==false)) > 0
        % only update if stats change
        set(handles.uitable_dataseries_assetselection,'Data',[data(:,1:2),newstatsdata]); % replace data in table with new values
    end
    
    % Return Series
    
    % Plot return series and add context menu for interactive deselection
    % get existing plot objectes
    plotobjects = get(handles.axes_dataseries_returnseries,'Children');
    if isempty(plotobjects)
        % Axes setup: 
        % use tight x axis limits
        % update x-axis date labels if needed 
        if ~isempty(dates)
            update_xaxis = true;
        else
            update_xaxis = false;
            dates = (0:size(returns,1))';  % dummy date vector
        end
        set(handles.axes_dataseries_returnseries,'XLim',[min(dates),max(dates)]);
        grid(handles.axes_dataseries_returnseries,'on')
        set(handles.axes_dataseries_returnseries,'Box','off');
        hold(handles.axes_dataseries_returnseries,'on');
        set(handles.axes_dataseries_returnseries,'Visible','on');
        title(handles.axes_dataseries_returnseries,'Return Series (right-click to deselect)');
    else
        % no need to update x axis (very slow...)
        update_xaxis = false;
        % remove all existing plot objects
        for i = 1:length(plotobjects)
            delete(plotobjects(i));
        end
    end
    % create custom colormap; even when plotting a subset of the return
    % series, each series always keeps its original color
    colormap = hsv(length(assetselection));
    if isempty(dates)
        dates = (0:size(returns,1))';  % dummy date vector for plotting
    end
    for i = 1:size(returns,2)
        po = plot(handles.axes_dataseries_returnseries,dates(2:end),returns(:,i),'Color',colormap(assetsel_ind(i),:));
        menu = uicontextmenu;  % create context menu
        item = uimenu(menu,'Label',['Deselect ''',assetlabels{i},''''],'Callback',@uimenu_callback);  % define callback
        set(item,'UserData',assetsel_ind(i));   % add asset position in table
        set(po,'UIContextMenu',menu);
    end
    % use centered y axis
    ylimmax = max(abs(returns(:)));
    set(handles.axes_dataseries_returnseries,'YLim',[-ylimmax,ylimmax]);
    % replace x-axis with date labels
    if update_xaxis
        datetick(handles.axes_dataseries_returnseries,'x','keeplimits');
    end

    % Results are no longer valid
    clearResultsPage(handles);
    
    % Enable controls
    set(handles.checkbox_dataseries_decayfactor,'Enable','on');
    set(handles.checkbox_dataseries_logreturns,'Enable','on');
    
    
        % Context menu callback function (nested)
        function uimenu_callback(menu,eventdata)
        % menu:       handle to context menu
        % eventdata:  not used
        
            % get selected asset
            assetlabel = get(menu,'UserData');
            % remove selected asset from list
            enableAsset(handles,assetlabel,false);
        end
end

% --- Update Portfolio Optimization Page
% Optimization panel
function updatePortfolioOptimizationPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % Only run if data available
    if isempty(handles.model.getPrices)
        return
    end

    % Enable controls
    set(handles.button_portopt_addconstraint,'Enable','on');
    set(handles.button_portopt_computeefficientfrontier,'Enable','on');

    % Get existing table content
    data = get(handles.uitable_portopt_genericconstraints,'Data');
    
    if isempty(data)
        % Allocate optimization page with new data 
        % after 'clearPortfolioOptimizationPage()', no constraints are
        % defined. We add constraint to sum up weights to 1 and limit all
        % assets to [0,1].
        labels = handles.model.getPricesLabels;
    
        % Constraint table
        set(handles.uitable_portopt_genericconstraints,'ColumnName',[labels(:)',{'Op','Value','Status'}]);
        set(handles.uitable_portopt_genericconstraints,'ColumnFormat',[repmat({[]},1,length(labels)),{{'=','<=','>='},'numeric',{'Enabled','Disabled','Select all','Deselect all','Delete'}}]);
        set(handles.uitable_portopt_genericconstraints,'ColumnEditable',true);
        set(handles.uitable_portopt_genericconstraints,'ColumnWidth',[repmat({80},1,length(labels)),{50,50,80}]);
        % Add one sample constraint (sum of weights = 1)
        button_portopt_addconstraint_Callback(handles.button_portopt_addconstraint,[],handles); 
    
        % Bound constraints table
        data = get(handles.uitable_portopt_boundconstraints,'Data');
        if isempty(data)
            % set default limits [0,1]
            set(handles.uitable_portopt_boundconstraints,'Data',[0;1]);
        end
    
    else
        % Update optimization page
        
        % get asset selection
        assetselection = handles.model.getAssetSelection;
        % set all unselected asset's column width to 0 and not editable
        columnwidth = [repmat(80,1,length(assetselection)),50,50,80];
        columnwidth(~assetselection) = 0;
        columnwidth = num2cell(columnwidth);
        set(handles.uitable_portopt_genericconstraints,'ColumnWidth',columnwidth);
        set(handles.uitable_portopt_genericconstraints,'ColumnEditable',[assetselection,true,true,true]);
    end
    
    
 
end

% --- Update Results Page
function updateResultsPage(handles)
% handles    structure with handles and user data (see GUIDATA)

    % First clear page to remove all prior allocation details
    clearResultsPage(handles);
    handles.selection = [];
    guidata(handles.figure, handles);

    % Enable controls
    set(handles.edit_results_confidencelevel,'Enable','on');
    set(handles.edit_results_riskfreerate,'Enable','on');
    set(handles.popupmenu_results_valueatrisk,'Enable','on');
    set(handles.button_results_createreport,'Enable','on');
    
    % Get statistics and optimization results
    % Note: use annualized values for visualization
    [~,~,~,annualized_ret,annualized_rsk] = handles.model.getStatistics;
    [~,~,annualized_benchmark_ret,annualized_benchmark_rsk] = handles.model.getBenchmarkStatistics;
    [~,~,pf_weights,annualized_pf_ret,annualized_pf_rsk] = handles.model.getOptimizationResults;
    % use percentages
    annualized_ret           = 100*annualized_ret;
    annualized_rsk           = 100*annualized_rsk;
    annualized_benchmark_ret = 100*annualized_benchmark_ret;
    annualized_benchmark_rsk = 100*annualized_benchmark_rsk;
    annualized_pf_ret        = 100*annualized_pf_ret;
    annualized_pf_rsk        = 100*annualized_pf_rsk;
    
    % Plot efficient frontier and individual assets
    axes(handles.axes_results_efficientfrontier);
    set(handles.axes_results_efficientfrontier,'Visible','on');
    hold('on');
    
%     plot(annualized_rsk,annualized_ret,'*','Color','r','MarkerSize',5);
    plot(annualized_pf_rsk,annualized_pf_ret,'-o','Color','b','MarkerSize',8);
    legend_str = {'Efficient Portfolios','Individual Assets'};
    if ~isempty(annualized_benchmark_rsk)
        plot(annualized_benchmark_rsk,annualized_benchmark_ret,'^','Color','k')
        legend_str = [legend_str,'Benchmark Portfoliio'];
    end
    grid('on');
    title('Select Portfolio');
    xlabel('Annualized Risk [%]');
    ylabel('Annualized Return [%]');
    % add legend, disable interactivity
    h = legend(legend_str,'Location','SouthEast');
    set(h,'UIContextMenu',[]);
    set(h,'HitTest','off');
    set(h,'Box','off')
    el = get(h,'Children');
    for i = 1:length(el)  %change text backgrounds to white
        if strcmp(get(el(i),'Type'),'text')
            set(el(i),'BackgroundColor',[1,1,1]);
        end
    end
    % save legend handle
    handles.axes_results_efficientfrontier_legend = h;
    guidata(handles.figure, handles);
    
    % Add marker and marker text to highlight individual assets (not visible right now)
    po = plot(annualized_pf_rsk(1),annualized_pf_ret(1),'s','MarkerSize',8,'Color','r','MarkerFaceColor','y');
    set(po,'Visible','off');
    set(po,'Userdata',-1);  % use -1 to find object later
    po = text(annualized_pf_rsk(1),annualized_pf_ret(1),'');
    set(po,'FontSize',8);
    set(po,'BackgroundColor',[1,1,1]);
    set(po,'EdgeColor',[0.8,0.8,0.8])
    set(po,'Visible','off');
    set(po,'Userdata',-2);  % use -2 to find object later
    
    

    
    % Add callback for interactive selection of individual assets
%     for i = 1:length(annualized_rsk)
%         po = plot(annualized_rsk(i),annualized_ret(i),'*','Color','r','MarkerSize',5);
%         set(po,'Userdata',i);  % asset number
%         set(po,'ButtonDownFcn',@assetselection_callback);
%     end
    % Add callback for interactive portfolio selection
    for i = 1:length(annualized_pf_rsk)
        po = plot(annualized_pf_rsk(i),annualized_pf_ret(i),'bo','MarkerFaceColor','w','MarkerSize',8);
        set(po,'Userdata',i);  % portfolio number
        set(po,'ButtonDownFcn',@portfolioselection_callback);
    end
        % Callback function (nested) for portfolio selection callback
        function portfolioselection_callback(hObject,event)
            % Select active portfolio
            if ~isempty(handles.selection)
                set(handles.selection,'MarkerFaceColor','none')
                set(handles.selection,'MarkerSize',8)
            end
            set(hObject,'MarkerFaceColor',[0,0.3,0.7])
            set(hObject,'MarkerSize',8)
            handles.selection = hObject; 
            guidata(handles.figure,handles);
            
            % Hide marker & markertext
            po = get(handles.axes_results_efficientfrontier,'Children');
            for j = 1:length(po)
                % marker has userdata "-1"
                if get(po(j),'Userdata') == -1
                    set(po(j),'Visible','off');
                end
                % markertext has userdata "-2"
                if get(po(j),'Userdata') == -2
                    set(po(j),'Visible','off');
                end
            end
            
            % Get selected portfolio number
            sel = get(handles.selection,'Userdata');
            
            % Update allocation chart
            weights = pf_weights(sel,:);
            labels  = handles.model.getPricesLabels;
            weights = weights(:);      % use column vectors
            labels  = labels(:);
            % collect all components with weight < 1% as "others"
            ind = abs(weights) > 0.02;
            alloc        = [weights(ind);sum(weights(~ind))];
            alloc_labels = [labels(ind);{'Others'}];
            % remove others if zero
            if abs(alloc(end)) < 1e-3
                alloc(end) = [];
                alloc_labels(end) = [];
            end
            % add weights to labels
            alloc_labels_weights = [];
            for j = 1:length(alloc_labels)
                alloc_labels_weights{j} = [alloc_labels{j},char(10),num2str(round(alloc(j)*10000)/100),'%'];
            end
            % Visualize allocation summary as pie chart (pie plot only accepts positive entries)
            axes(handles.axes_results_allocation);
            h = pie(abs(alloc),alloc_labels_weights);
            for j = 1:length(h)
                if strcmp(get(h(j),'Type'),'text')
                    set(h(j),'FontSize',7);
                end
                if strcmp(get(h(j),'Type'),'patch')
                    set(h(j),'FaceAlpha',0.7);
                    set(h(j),'EdgeAlpha',0.2);
                end
            end

            % Show weights details in table
            ind = abs(weights) > 0.0001;
            alloc        = [weights(ind)];
            alloc_labels = [labels(ind)];            
            [alloc,ind] = sort(alloc,'descend');  % sort alloc and labels in descending order
            alloc_labels = alloc_labels(ind);
            data = {};
            for j = 1:length(alloc)
                data = [data;[alloc_labels{j},'  (',sprintf('%2.2f',alloc(j)*100),'%)']];
            end
            set(handles.uitable_results_weights,'Data',data);
            % Set column width
            pos = get(handles.uitable_results_weights,'Position');
            if length(alloc) > 14


                tablewidth = pos(3) - 4 - 16;  % border + slider
            else
                tablewidth = pos(3) - 4;  % border only
            end
            set(handles.uitable_results_weights,'ColumnWidth',num2cell(tablewidth));


            % Visualize portfolio performance, compare to benchmark
            axes(handles.axes_results_performance);
            set(handles.axes_results_performance,'Visible','on');
            hold('off')
            dates = handles.model.getDates;
            prices = handles.model.getPrices;
            pf_prices = prices*weights;
            pf_prices = 100*pf_prices/abs(pf_prices(1));  % normalize
            if pf_prices(1) < 0
                pf_prices = pf_prices + 200;   % shift up if first val = -100
            end
            legend_str = 'Selected Portfolio';
            if ~isempty(dates)
                plot(dates,pf_prices);
                axis('tight');
                datetick('x','keeplimits');
            else
                plot(pf_prices);
            end
            grid('on');
            box('off');
            ylabel('Relative Performance [%]');
            hold('on');
            benchmark = handles.model.getBenchmark;
            if ~isempty(benchmark)
                legend_str = {legend_str,handles.model.getBenchmarkLabel};
                benchmark = 100*benchmark/benchmark(1);   % normalize
                if ~isempty(dates)
                    plot(dates,benchmark,'r');
                    axis('tight');
                    datetick('x','keeplimits');
                else
                    plot(benchmark,'r');
                end
            end
            h = legend(legend_str,'Location','NorthWest');
            set(h,'UIContextMenu',[]);
            set(h,'HitTest','off');
            set(h,'Box','off')
            el = get(h,'Children');
            for j = 1:length(el)  %change text backgrounds to white
                if strcmp(get(el(j),'Type'),'text')
                    set(el(j),'BackgroundColor',[1,1,1]);
                end
            end
            % save legend handle
            handles.axes_results_performance_legend = h;
            guidata(hObject, handles);

            % Update performance metrics (fire event to riskfree rate edit box)
            edit_results_riskfreerate_Callback(handles.edit_results_riskfreerate, [], handles);
            
            % Update value at risk (fire event to confidence level edit box)
            set(handles.axes_results_valueatrisk,'Visible','on');
            edit_results_confidencelevel_Callback(handles.edit_results_confidencelevel, [], handles);
            
        end
    
        % Callback function (nested) for asset selection callback
        function assetselection_callback(hObject,event)

            % get selected item
            selection = get(hObject,'Userdata');

            % update marker
            showMarker(handles,selection)

        end
end


% --- Make selected tab and panel active (visible).
function tab_handler(active_tab,handles)
    % activate selected tab and disable all others
    for i = 1:handles.tabcount
        h = eval(['handles.axes_tab_',num2str(i)]);
        p = findobj(h,'Type','patch');
        if ~isempty(p)
            if i==active_tab
                set(p,'FaceColor',[1,1,1]);
                eval(['set(handles.uipanel_tab',num2str(i),',''Visible'',''on'');']);
            else
                %set(p,'FaceColor',[0.9,0.95,1]);
                set(p,'FaceColor',[0.8,0.85,1]);
                eval(['set(handles.uipanel_tab',num2str(i),',''Visible'',''off'');']);
            end
        end
    end
end


% --- Executes on mouse press over axes background.
function axes_tab_1_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to axes_tab_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
    tab_handler(1,handles);
end


% --- Executes on mouse press over axes background.
function axes_tab_2_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to axes_tab_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
    tab_handler(2,handles);
end


% --- Executes on mouse press over axes background.
function axes_tab_3_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to axes_tab_3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
    tab_handler(3,handles);
end


% --- Executes on mouse press over axes background.
function axes_tab_4_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to axes_tab_4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
    tab_handler(4,handles);
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_selectsource_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_selectsource (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


% --- Executes on selection change in popupmenu_dataimport_selectsource.
function popupmenu_dataimport_selectsource_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_selectsource (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % get selection
    % legend is
    % 1: MATLAB Base Workspace
    % 2: Yahoo Datafeed
    % 3: Excel
    selection = get(handles.popupmenu_dataimport_selectsource,'Value');
    % Activate selected import panel and reset fields
    switch selection
        case 1  % MATLAB Workspace
            if strcmp(get(handles.uipanel_dataimport_workspace,'Visible'),'off')
                % Activate panel
                set(handles.uipanel_dataimport_workspace,'Visible','on');
                set(handles.uipanel_dataimport_datafeed,'Visible','off');
                set(handles.uipanel_dataimport_xlsfile,'Visible','off');
                % Read workspace variables and assign them to popup menus
                vars = evalin('base', 'whos');
                var_list_numeric = {};
                var_list_cellstrings = {};
                var_list_strings = {};
                for i = 1:length(vars)
                    if strcmp(vars(i).class, 'double')
                        var_list_numeric = [var_list_numeric; vars(i).name]; 
                    end
                    if strcmp(vars(i).class, 'cell')
                        var_list_cellstrings = [var_list_cellstrings; vars(i).name]; 
                    end
                    if strcmp(vars(i).class, 'char')
                        var_list_strings = [var_list_strings; vars(i).name]; 
                    end
                end
                var_list_dates = sort([var_list_cellstrings;var_list_numeric]);  % dates may be strings or serial
                if ~isempty(var_list_numeric)
                    var_list_numeric = ['Select variable'; var_list_numeric];
                else
                    var_list_numeric = 'No data available';
                end
                if ~isempty(var_list_cellstrings)
                    var_list_cellstrings = ['Select variable'; var_list_cellstrings];
                else
                    var_list_cellstrings = 'No data available';
                end
                if ~isempty(var_list_strings)
                    var_list_strings = ['Select variable'; var_list_strings];
                else
                    var_list_strings = 'No data available';
                end
                if ~isempty(var_list_dates)
                    
                    var_list_dates = ['Select variable'; var_list_dates];
                else
                    var_list_dates = 'No data available';
                end
                % assign variables list to all popup menus
                set(handles.popupmenu_dataimport_workspace_prices,'String',var_list_numeric)
                set(handles.popupmenu_dataimport_workspace_benchmark,'String',var_list_numeric)
                set(handles.popupmenu_dataimport_workspace_dates,'String',var_list_dates)  
                set(handles.popupmenu_dataimport_workspace_priceslabels,'String',var_list_cellstrings)
                set(handles.popupmenu_dataimport_workspace_benchmarklabel,'String',var_list_strings)
            end
        case 2  % Yahoo Datafeed
            if strcmp(get(handles.uipanel_dataimport_datafeed,'Visible'),'off')
                % Activate panel
                set(handles.uipanel_dataimport_datafeed,'Visible','on');
                set(handles.uipanel_dataimport_workspace,'Visible','off');
                set(handles.uipanel_dataimport_xlsfile,'Visible','off');
            end
        case 3  % Excel file
            if strcmp(get(handles.uipanel_dataimport_xlsfile,'Visible'),'off')
                % Activate panel
                set(handles.uipanel_dataimport_xlsfile,'Visible','on');
                set(handles.uipanel_dataimport_workspace,'Visible','off');
                set(handles.uipanel_dataimport_datafeed,'Visible','off');
            end
    end

end


% --- Executes on selection change in popupmenu_dataimport_workspace_prices.
function popupmenu_dataimport_workspace_prices_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_prices (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_dataimport_workspace_prices contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_dataimport_workspace_prices
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_workspace_prices_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_prices (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end

% --- Executes on selection change in popupmenu_dataimport_workspace_benchmark.
function popupmenu_dataimport_workspace_benchmark_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_benchmark (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_dataimport_workspace_benchmark contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_dataimport_workspace_benchmark
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_workspace_benchmark_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_benchmark (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end

% --- Executes on selection change in popupmenu_dataimport_workspace_dates.
function popupmenu_dataimport_workspace_dates_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_dates (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_dataimport_workspace_dates contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_dataimport_workspace_dates
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_workspace_dates_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_dates (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end

% --- Executes on button press in button_dataimport_workspace_importdata.
function button_dataimport_workspace_importdata_Callback(hObject, eventdata, handles)
% hObject    handle to button_dataimport_workspace_importdata (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Import Workspace data
    
    % Reset controls
    set(handles.text_dataimport_workspace_status,'String','');
    disableImportDataPage(handles);
    
    % Prices
    selection = get(handles.popupmenu_dataimport_workspace_prices,'Value');
    if selection == 1
        set(handles.text_dataimport_workspace_status,'String','Select prices series');
        return
    end
    variable = get(handles.popupmenu_dataimport_workspace_prices,'String');
    variable = variable{selection};
    prices = evalin('base', variable);
    if size(prices,1) < 2 || size(prices,2) < 2
        set(handles.text_dataimport_workspace_status,'String','Prices variable must have at least 2 columns');
        return
    end
    % Benchmark
    selection = get(handles.popupmenu_dataimport_workspace_benchmark,'Value');
    if selection == 1
        % no benchmark
        benchmark = [];
    else
        variable = get(handles.popupmenu_dataimport_workspace_benchmark,'String');
        variable = variable{selection};
        benchmark = evalin('base', variable);
        if ~isvector(benchmark) || length(benchmark) ~= size(prices,1)
            set(handles.text_dataimport_workspace_status,'String',{'Benchmark series must be a vector with same number';'of elements as in each prices series'});
            return
        end
    end
    % Dates
    selection = get(handles.popupmenu_dataimport_workspace_dates,'Value');
    if selection == 1
        % no dates
        dates = [];
    else
        variable = get(handles.popupmenu_dataimport_workspace_dates,'String');
        variable = variable{selection};
        dates = evalin('base', variable);
        if ~isvector(dates) || length(dates) ~= size(prices,1)
            set(handles.text_dataimport_workspace_status,'String',{'Dates series must be a vector with same number';'of elements as in each prices series'});
            return
        end
        % Read date string format
        selection = get(handles.popupmenu_dataimport_workspace_datestringformat,'Value');
        options   = get(handles.popupmenu_dataimport_workspace_datestringformat,'String');
        datestringformat = options{selection};
        % convert to serial dates
        switch datestringformat
            case 'Serial date number (MATLAB)'
            case 'Serial date number (Excel)'
                % convert to MATLAB serial dates
                dates = x2mdate(dates);
            otherwise
                try
                    dates = datenum(dates,datestringformat);
                catch
                    set(handles.text_dataimport_workspace_status,'String',{'Date string format does not match'});
                    return
                end
                % save date string format persistently
                handles.datestringformat = datestringformat;
                guidata(hObject, handles);
        end
    end
    % Prices labels
    selection = get(handles.popupmenu_dataimport_workspace_priceslabels,'Value');
    if selection == 1
        % no prices labels, create default labels
        num = size(prices,2);  
        labels = cellstr([repmat('Asset ',num,1),num2str((1:num)')]);
        priceslabels = strrep(labels,'  ',' ');
    else
        variable = get(handles.popupmenu_dataimport_workspace_priceslabels,'String');
        variable = variable{selection};
        priceslabels = evalin('base', variable);
        if ~isvector(priceslabels) || length(priceslabels) ~= size(prices,2)
            set(handles.text_dataimport_workspace_status,'String',{'Prices labels series must be a vector with same number';'of elements as prices series'});
            return
        end
    end
    % Benchmark label
    if isempty(benchmark)
        benchmarklabel = [];
    else
        selection = get(handles.popupmenu_dataimport_workspace_benchmarklabel,'Value');
        if selection == 1
            % no benchmark label, create default
            benchmarklabel = 'Benchmark Index';
        else
            variable = get(handles.popupmenu_dataimport_workspace_benchmarklabel,'String');
            variable = variable{selection};
            benchmarklabel = evalin('base', variable);
            if ~ischar(benchmarklabel)
                set(handles.text_dataimport_workspace_status,'String','Benchmark label series must be a string');
                return
            end
        end
    end
    
    % Write data to table
    updateImportDataPage(handles,prices,benchmark,dates,priceslabels,benchmarklabel);

    
end


% --- Executes on selection change in popupmenu_dataimport_workspace_priceslabels.
function popupmenu_dataimport_workspace_priceslabels_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_priceslabels (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_dataimport_workspace_priceslabels contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_dataimport_workspace_priceslabels
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_workspace_priceslabels_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_priceslabels (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end

% --- Executes on selection change in popupmenu_dataimport_workspace_benchmarklabel.
function popupmenu_dataimport_workspace_benchmarklabel_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_benchmarklabel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_dataimport_workspace_benchmarklabel contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_dataimport_workspace_benchmarklabel
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_workspace_benchmarklabel_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_benchmarklabel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


% --- Executes on button press in button_portopt_addconstraint.
function button_portopt_addconstraint_Callback(hObject, eventdata, handles)
% hObject    handle to button_portopt_addconstraint (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Only run if data available
    if isempty(handles.model.getPrices)
        return
    end
    
    % add an additional row to constraints table
    % if it is first row, all assets are enabled
    % otherwise all assets are disabled
    data = get(handles.uitable_portopt_genericconstraints,'Data');
    % get total number of assets
    numassets = length(handles.model.getAssetSelection);
    if isempty(data)
        data = [repmat({true},1,numassets),{'=',1,'Enabled'}];
    else
        data = [data;repmat({false},1,numassets),{'=',1,'Enabled'}];
    end
    set(handles.uitable_portopt_genericconstraints,'Data',data);

end

% --- Executes on button press in button_portopt_computeefficientfrontier.
function button_portopt_computeefficientfrontier_Callback(hObject, eventdata, handles)
% hObject    handle to button_portopt_computeefficientfrontier (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Only run if data available
    if isempty(handles.model.getPrices)
        return
    end

    % disable control
    set(handles.button_portopt_computeefficientfrontier,'Enable','off');
    drawnow('expose');
    
    % clear results page
    clearResultsPage(handles);
    
    % Create constraints set
    conSet = [];

    % Asset selection
    assetselection = handles.model.getAssetSelection;
    numAssets = sum(assetselection);

    % Constraint table
    % get data
    data = get(handles.uitable_portopt_genericconstraints,'Data');
    % only use selected assets + options fields
    data = data(:,[assetselection,true,true,true]);  
    % remove rows with status 'Disabled'
    badrows = strcmp(data(:,end),'Disabled');
    data(badrows,:) = [];
    % remove rows with no selected asset
    badrows = ~logical(sum(cell2mat(data(:,1:end-3)),2));
    data(badrows,:) = [];
    for i = 1:size(data,1)
        switch data{i,numAssets+1}
            case '<='
                conSet = [conSet;
                          cell2mat(data(i,1:numAssets)),data{i,numAssets+2}];
            case '>='
                conSet = [conSet;
                          -cell2mat(data(i,1:numAssets)),-data{i,numAssets+2}];
            case '='
                conSet = [conSet;
                          cell2mat(data(i,1:numAssets)),data{i,numAssets+2};
                          -cell2mat(data(i,1:numAssets)),-data{i,numAssets+2}];
        end
    end
    % Asset limits
    data = get(handles.uitable_portopt_boundconstraints,'Data');   
    if isnan(data(1))
        data(1) = 0;
        set(handles.uitable_portopt_boundconstraints,'Data',data);
    end
    if isnan(data(2))
        data(2) = 1;
        set(handles.uitable_portopt_boundconstraints,'Data',data);
    end
    if data(2) < data(1)
        data = [0;1];
        set(handles.uitable_portopt_boundconstraints,'Data',data);
    end
    conSet = [conSet;
              -eye(numAssets),-ones(numAssets,1)*data(1);
              eye(numAssets),ones(numAssets,1)*data(2)];
          
    
    % Optimize portfolios 
    msg = handles.model.computeEfficientFrontier(conSet);
    if ~isempty(msg)
        % enable control
        set(handles.button_portopt_computeefficientfrontier,'Enable','on');    
        msgbox(msg,'Warning');
        return
    end
    
    % Update results page
    updateResultsPage(handles);
    
    % Make results page current
    tab_handler(3,handles);

    % enable control
    set(handles.button_portopt_computeefficientfrontier,'Enable','on');    
end


% --- Executes on button press in checkbox_dataseries_logreturns.
function checkbox_dataseries_logreturns_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_dataseries_logreturns (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Apply selection to model
    val = boolean(get(handles.checkbox_dataseries_logreturns,'Value'));
    handles.model.useLogReturns(val);
    
    % Update visualization
    updateDataSeriesPage(handles)
end


% --- Executes on selection change in popupmenu_dataimport_workspace_datestringformat.
function popupmenu_dataimport_workspace_datestringformat_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_datestringformat (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_dataimport_workspace_datestringformat contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_dataimport_workspace_datestringformat
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_workspace_datestringformat_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_workspace_datestringformat (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


% --- Executes when entered data in editable cell(s) in uitable_portopt_genericconstraints.
function uitable_portopt_genericconstraints_CellEditCallback(hObject, eventdata, handles)
% hObject    handle to uitable_portopt_genericconstraints (see GCBO)
% eventdata  structure with the following fields (see UITABLE)
%	Indices: row and column indices of the cell(s) edited
%	PreviousData: previous data for the cell(s) edited
%	EditData: string(s) entered by the user
%	NewData: EditData or its converted form set on the Data property. Empty if Data was not changed
%	Error: error string when failed to convert EditData to appropriate value for Data
% handles    structure with handles and user data (see GUIDATA)

    % Get table data
    data = get(handles.uitable_portopt_genericconstraints,'Data');
    % Only proceed if last column is affected
    if eventdata.Indices(2) ~= size(data,2)
        return
    end
    do_update = false;

    % If action "Delete" has been selected, remove complete row from table
    if strcmp(eventdata.NewData,'Delete')
       data(eventdata.Indices(1),:) = []; 
       do_update = true;
    end

    % If action "Select all / Unselect all" selected,
    % update checkboxes and change status back to old value
    ind = ~cellfun('isempty',strfind({'Select all','Deselect all'},eventdata.NewData));
    if any(ind)
       data(eventdata.Indices(1),1:end-3) = {ind(1)}; 
       data(eventdata.Indices(1),end) = {eventdata.PreviousData}; 
       do_update = true;
    end
    
    % Update table if needed
    if do_update
        set(handles.uitable_portopt_genericconstraints,'Data',data);
    end

end


% --- Executes when selected cell(s) is changed in uitable_results_weights.
function uitable_results_weights_CellSelectionCallback(hObject, eventdata, handles)
% hObject    handle to uitable_results_weights (see GCBO)
% eventdata  structure with the following fields (see UITABLE)
%	Indices: row and column indices of the cell(s) currently selecteds
% handles    structure with handles and user data (see GUIDATA)

    % get selected item
    if isempty(eventdata) || isempty(eventdata.Indices)
        return
    end
    selection = eventdata.Indices;
    selection = selection(1);  % only need row index
    sel_label = get(handles.uitable_results_weights,'Data');
    sel_label = sel_label{selection,1};  % extract selected asset label
    ind1 = find(sel_label=='(',1,'first');
    if isempty(ind1)
        return
    end
    sel_label = strtrim(sel_label(1:ind1-1));
    
    % update marker
    showMarker(handles,sel_label)
    
end

% --- Helper function to update marker on efficient frontier
function showMarker(handles,selection)
% selection    Selected asset number or asset label

    % get marker/markertext handle
    po = get(handles.axes_results_efficientfrontier,'Children');
    marker = [];
    markertext = [];
    for i = 1:length(po)
        % marker has userdata "-1"
        if get(po(i),'Userdata') == -1
            marker = po(i);
        end
        % markertext has userdata "-2"
        if get(po(i),'Userdata') == -2
            markertext = po(i);
        end
    end
    if isempty(marker) || isempty(markertext)
        return
    end
    
    % get statistics and extract corresponding asset
    [~,~,~,annualized_ret,annualized_rsk] = handles.model.getStatistics;
    % use percentages
    annualized_ret           = 100*annualized_ret;
    annualized_rsk           = 100*annualized_rsk;
    labels      = handles.model.getPricesLabels;
    if ischar(selection)
        ind = find(strcmp(labels,selection));
    else
        ind = selection;
    end
    
    % update marker / markertext
    if ~isempty(ind) && isscalar(ind)
        % update marker data
        set(marker,'XData',annualized_rsk(ind),'YData',annualized_ret(ind)); 
        set(marker,'Visible','on');
        % update marker text
        set(markertext,'String',labels{ind});
        % find good position
        m_extend = get(markertext,'Extent');
        a_xlim = get(handles.axes_results_efficientfrontier,'XLim');
        final_pos = zeros(1,2);  % place text to the right of marker
        final_pos(1) = annualized_rsk(ind) + range(a_xlim)/40; 
        final_pos(2) = annualized_ret(ind);
        if a_xlim(2) - (final_pos(1) + m_extend(3)) < 0
            % not enough space, move it to left side
            final_pos(1) = annualized_rsk(ind) - range(a_xlim)/40 - m_extend(3);
        end
        set(markertext,'Position',final_pos);
        set(markertext,'Visible','on');
    else
        set(marker,'Visible','off');
        set(markertext,'Visible','off');
    end    
end


function edit_results_riskfreerate_Callback(hObject, eventdata, handles)
% hObject    handle to edit_results_riskfreerate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_results_riskfreerate as text
%        str2double(get(hObject,'String')) returns contents of edit_results_riskfreerate as a double

    % Update risk metrics
    
    % get portfolio number
    selection = get(handles.selection,'Userdata');  
    if isempty(selection)
        return
    end

    % get risk-free rate
    riskfreerate = str2double(get(handles.edit_results_riskfreerate,'String'));
    if isnan(riskfreerate)
        riskfreerate = 2;
        set(handles.edit_results_riskfreerate,'String','2');
    end
    
    % get key metrics
    metrics = handles.model.getPerformanceMetrics(selection, riskfreerate/100);
    if isempty(metrics)
        return
    end
    
    % collect data
    % some fiels are not useful without benchmark index
    if isempty(handles.model.getBenchmark)
        data = {'Annualized Volatility',[sprintf('%2.2f',100*metrics.annualizedvolatility),'%']; ...
                'Annualized Return',[sprintf('%2.2f',100*metrics.annualizedreturn),'%']; ...
                'Correlation','-'; ...
                'Sharpe Ratio',sprintf('%2.2f',metrics.sharperatio); ...
                'Alpha','-'; ...
                'Risk-adjusted Return','-'; ...
                'Information Ratio','-'; ...
                'Tracking Error','-'; ...
                'Max. Drawdown',[sprintf('%2.2f',100*metrics.maxdrawdown),'%']};
    else
        data = {'Annualized Volatility',[sprintf('%2.2f',100*metrics.annualizedvolatility),'%']; ...
                'Annualized Return',[sprintf('%2.2f',100*metrics.annualizedreturn),'%']; ...
                'Correlation',sprintf('%2.2f',metrics.correlation); ...
                'Sharpe Ratio',sprintf('%2.2f',metrics.sharperatio); ...
                'Alpha',[sprintf('%2.2f',100*metrics.alpha),'%']; ...
                'Risk-adjusted Return',[sprintf('%2.2f',100*metrics.riskadjusted_return),'%']; ...
                'Information Ratio',[sprintf('%2.2f',100*metrics.inforatio),'%']; ...
                'Tracking Error',[sprintf('%2.2f',100*metrics.trackingerror),'%']; ...
                'Max. Drawdown',[sprintf('%2.2f',100*metrics.maxdrawdown),'%']};
    end    
    
    % Add all metrics to table
    set(handles.uitable_results_metrics,'Data',data,'ColumnName',[]);
    
end

% --- Executes during object creation, after setting all properties.
function edit_results_riskfreerate_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_results_riskfreerate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


function edit_results_confidencelevel_Callback(hObject, eventdata, handles)
% hObject    handle to edit_results_confidencelevel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Update Value at Risk
    
    % get portfolio number
    selection = get(handles.selection,'Userdata');  
    if isempty(selection)
        return
    end

    % get confidence level
    confidence_level = str2double(get(handles.edit_results_confidencelevel,'String'));
    if isnan(confidence_level)
        confidence_level = 95;
        set(handles.edit_results_confidencelevel,'String','95');
    end

    % Compute VaR depending on selected option
    option = get(handles.popupmenu_results_valueatrisk,'Value');
    if option == 1
        handles.model.computeHistoricalVaR(selection,confidence_level/100,handles.axes_results_valueatrisk);
    else
        handles.model.computeParameticVaR(selection,confidence_level/100,handles.axes_results_valueatrisk);
    end

end

% --- Executes during object creation, after setting all properties.
function edit_results_confidencelevel_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_results_confidencelevel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


% --- Executes on selection change in popupmenu_results_valueatrisk.
function popupmenu_results_valueatrisk_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_results_valueatrisk (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Trigger update (fire event to confidence level edit box)
    edit_results_confidencelevel_Callback(handles.edit_results_confidencelevel, [], handles);
    
end 

% --- Executes during object creation, after setting all properties.
function popupmenu_results_valueatrisk_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_results_valueatrisk (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


% --- Executes on button press in button_results_createreport.
function button_results_createreport_Callback(hObject, eventdata, handles)
% hObject    handle to button_results_createreport (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Options
    DisableReporting = false;  % set to true to disable reporting
    ExcelSheetVisible = true;  % set to true to keep Excel sheet visible during reporting process
    FileName = 'Report.xlsx';
    SheetName = 'Summary';
    Overwrite = true;
    
    
    % Delete existing report if necessary
    if Overwrite && exist(FileName,'file')
        delete(FileName);
    end
    
    % create formatting templates
    % text normal
    opts_txt_normal = ExcelReport.getDefaultTextOptions;
    opts_txt_normal.FontSize = 11;
    % text title 1
    opts_txt_title1 = ExcelReport.getDefaultTextOptions;
    opts_txt_title1.Bold = true;
    opts_txt_title1.FGColor = [0.15,0.15,0.15];
    opts_txt_title1.FontSize = 20;
    % text title 2
    opts_txt_title2 = opts_txt_title1;
    opts_txt_title2.FontSize = 14;
    % text title 3
    opts_txt_title3 = opts_txt_title1;
    opts_txt_title3.FontSize = 12;
    % data table
    opts_table_data = ExcelReport.getDefaultTableOptions;
    opts_table_data.MajorGrid = true;
    opts_table_data.MinorGrid = true;
    opts_table_data.GridColor = [0.2,0.2,0.2];
    opts_table_data.HorizontalAlignment = 'center';
    opts_table_data.Title.Grid = true;
    opts_table_data.Title.GridColor = [0.2,0.2,0.2];
    
    % Setup report
    % ------------
    report = ExcelReport(FileName,SheetName,DisableReporting,ExcelSheetVisible);
    report.setOrientation('Landscape')
    report.setColumnWidth('A:ZZ',3.1);
    report.setRowHeight('1:1000',15);
    row_count = 1;

    % Title / Header
    % --------------
    report.insertText([row_count,1],'Portfolio Allocation Report',opts_txt_title1);
    [~,s] = weekday(today,'long');
    report.insertText([row_count + 1,1],[' ',s,', ',datestr(today,handles.datestringformat)],opts_txt_normal);
    % add logo
    report.insertPicture([row_count,50],'membrane_medium_col.eps',[137,120]);
    row_count = row_count + 3;
    
    % Optimzation setup
    % -----------------
    report.insertText([row_count,1],'Optimization Setup',opts_txt_title2);
    row_count = row_count + 2;

    % Limit constraints
    data = get(handles.uitable_portopt_boundconstraints,'Data');   
    report.insertText([row_count,1],'Asset Limits',opts_txt_title3);
    report.insertText([row_count+2,1],['Lower limit:  ',sprintf('%2i',data(1)*100),'%'],opts_txt_normal);
    report.insertText([row_count+3,1],['Upper limit:  ',sprintf('%2i',data(2)*100),'%'],opts_txt_normal);
    report.setRowHeight(num2str(row_count+1),5);
    
    % Generic constraints
    data = get(handles.uitable_portopt_genericconstraints,'Data');
    % only use selected assets + options fields
    assetselection = handles.model.getAssetSelection;
    data = data(:,[assetselection,true,true,true]);  
    % remove rows with status 'Disabled'
    badrows = strcmp(data(:,end),'Disabled');
    data(badrows,:) = [];
    % remove rows with no selected asset
    badrows = ~logical(sum(cell2mat(data(:,1:end-3)),2));
    data(badrows,:) = [];
    % get labels
    priceslabels = handles.model.getPricesLabels();
    % collect constraints
    constraints = {};
    for row = 1:size(data,1)
        % get active labels
        ind = cell2mat(data(row,1:end-3));
        % create list of labels
        if all(ind)
            txt = 'Sum of all assets';
        else
            labels = priceslabels(cell2mat(data(row,1:end-3)));
            txt = labels{1};
            for count = 2:length(labels)
                % break line after 10 items
                if mod(count,10) == 1
                    % add expression to list
                    constraints{end+1} = txt;
                    txt = '';
                end
                txt = [txt,'  +  ',labels{count}];
            end
        end
        % add expression
        txt = [txt,'  ',data{row,end-2},'  ',[sprintf('%2i',data{row,end-1}*100),'%']];
        % add to list
        constraints{end+1} = txt;
    end
    % only continue if any data available
    if ~isempty(constraints)
        % optimzation constraints
        report.insertText([row_count,9],'Constraints',opts_txt_title3);
        for el = 1:length(constraints)
            report.insertText([row_count+1+el,9],constraints{el},opts_txt_normal);
        end
        % increase row height
        report.setRowHeight([num2str(row_count+2),':',num2str(row_count+1+length(constraints))],17);
    end
     
    % compute row count
    row_count = row_count + max(5,length(constraints)+4);
    
    % Optimization results
    % --------------------
    report.insertText([row_count,1],'Optimization Results',opts_txt_title2);
    row_count = row_count + 2;
    % efficient frontier (including legend)
    report.insertFigure([row_count,1],[handles.axes_results_efficientfrontier,handles.axes_results_efficientfrontier_legend])
    % portfolio summary
    report.insertText([row_count,20],'Portfolio Summary',opts_txt_title3);
    % get optim results from model
    [~, pf_rsk, pf_weights,annualizedreturn,annualizedvolatility] = handles.model.getOptimizationResults();
    priceslabels = handles.model.getPricesLabels();
    % build nicely formatted table content
    pf_weights = pf_weights';
    skip = sum(abs(pf_weights),2) < 0.01; % do not display unused assets
    pf_weights(skip,:) = [];
    priceslabels(skip) = [];
    data = [1:length(pf_rsk);annualizedvolatility';annualizedreturn';pf_weights];
    data = num2cell(data);
    for row = 1:size(data,1)
        for col = 1:size(data,2)
            if row == 1
                data{row,col} = num2str(data{row,col});
            else
                if abs(data{row,col}) < 0.01
                    data{row,col} = '-';
                else
                    data{row,col} = [sprintf('%2.2f',data{row,col}*100),'%'];
                    if strcmp(data{row,col}(end-3:end),'.00%')  % remove trailing zeros
                        data{row,col}(end-3:end-1) = [];
                    end
                end
            end
        end
    end
    data = [[{'Portfolio #';'Annualized Volatility';'Annualized Return'};priceslabels(:)],data];
    opts_table_data.MergeNumberOfColumns = [6,2*ones(1,size(data,2)-1)];
    opts_table_data.FontSize = 9;
    % insert data into Excel, use two headers to separate performance and weights
    report.insertTable([row_count+2,20],data(1:3,:),'Performance',opts_table_data);
    report.insertTable([row_count+6,20],data(4:end,:),'Weights',opts_table_data);
    
    % graph needs at least 15 rows
    row_count = row_count + max(15,size(data,1)+5);
    
    
    % Allocation and metrics of selected portfolio
    % --------------------------------------------
    % run only if a portfolio selected on efficient frontier
    if ~isempty(handles.selection)
        % Get selected portfolio number
        sel = get(handles.selection,'Userdata');
        % portfolio allocation
        report.insertText([row_count,1],['Selected Portfolio #',num2str(sel)],opts_txt_title2);
        row_count = row_count + 2;
        % portfolio allocation
        report.insertText([row_count,1],'Allocation',opts_txt_title3);
        % separate data from table into two columns
        data = get(handles.uitable_results_weights,'Data'); 
        for i = 1:size(data,1)
            ind1 = find(data{i,1}=='(',1,'last');
            ind2 = find(data{i,1}==')',1,'last');
            if ~isempty(ind1) && ~isempty(ind2)
                data{i,2} = data{i,1}(ind1+1:ind2-1);
                data{i,1} = strtrim(data{i,1}(1:ind1-1));
            end
        end
        opts_table_data.MergeNumberOfColumns = [7,2];
        report.insertTable([row_count + 2,1],data,[],opts_table_data);
        % allocation chart
        report.insertFigure([row_count + 2,11],handles.axes_results_allocation)
 
        % performance
        report.insertText([row_count,23],'Performance',opts_txt_title3);
        % performance chart (including legend)
        report.insertFigure([row_count + 2,23],[handles.axes_results_performance,handles.axes_results_performance_legend],[350,220])
        
        % key metrics
        report.insertText([row_count,38],'Key Metrics',opts_txt_title3);
        % metrics
        data = get(handles.uitable_results_metrics,'Data');
        opts_table_data.MergeNumberOfColumns = [5,2];
        report.insertTable([row_count + 3,38],data,[],opts_table_data);
        % VaR
        report.insertFigure([row_count + 1,46],handles.axes_results_valueatrisk,[280,260]);
        
    end
    
    % Make sure report fits on one page
    report.setFitToPage();
    
    % Create pdf
    report.createPDF('Report.pdf');

    % Finalize Excel report
    report.closeReport();
    
    % Open final pdf
    winopen('Report.pdf')

end


% --- Executes on selection change in popupmenu_dataimport_datafeed_indexname.
function popupmenu_dataimport_datafeed_indexname_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_datafeed_indexname (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu_dataimport_datafeed_indexname contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu_dataimport_datafeed_indexname
end

% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_datafeed_indexname_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_datafeed_indexname (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end

% --- Executes on button press in button_dataimport_datafeed_fetchsymbols.
function button_dataimport_datafeed_fetchsymbols_Callback(hObject, eventdata, handles)
% hObject    handle to button_dataimport_datafeed_fetchsymbols (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Fetch Symbols of selected Index components
    
    % Hide download panel
    set(handles.uipanel_dataimport_datafeed_download,'Visible','off');
    
    % Disable 'fetch symbol' button
    set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','off');
    % update status
    set(handles.text_dataimport_datafeed_symbollookup_status,'String','Download in progress.')
    
    % get index name from popup menu
    % hint: http://finance.yahoo.com/intlindices lists all available indices
    indexname = cellstr(get(handles.popupmenu_dataimport_datafeed_indexname,'String'));
    selection = get(handles.popupmenu_dataimport_datafeed_indexname,'Value');
    indexname = indexname{selection};
    ind = find(indexname=='(',1,'first'); % remove description in brackets
    if ~isempty(ind)
        indexdesc = strtrim(indexname(ind+1:end-1));
        indexname = strtrim(indexname(1:ind-1));
    end
    
    % fetch symbols from Yahoo! Finance symbol lookup page (multiple pages)
    try 
        h = actxserver('internetexplorer.application');
        h.Visible = 0;

        symbols = {};
        pagecount = 0;
        done = false;
        while ~done
            % navigate to symbol lookup page
            h.Navigate(['http://finance.yahoo.com/q/cp?s=',indexname,'&c=',num2str(pagecount)]);
            % wait for frameset (15 seconds max)
            pause(0.2);
            entry = now;
            while ((h.Busy ~= 0) || ~strcmp(h.readyState,'READYSTATE_COMPLETE')) && (now-entry)*24*60*60 < 15
                pause(0.2)
            end
            
            % update status (add '.')
            txt = get(handles.text_dataimport_datafeed_symbollookup_status,'String');
            set(handles.text_dataimport_datafeed_symbollookup_status,'String',[txt,'.']);

            % break while loop if no table found below
            done = true;
           
            % get all tables
            tables = h.document.querySelectorAll('table'); 
%             tables = h.document.getElementsByTagName('table');
            for t = 0:tables.length-1
                % scan rows if current table has at least 10 cells
                if tables.item(t).cells.length >= 10
                    % continue with next page
                    done = false;
                    % get row elements
                    for r = 0:tables.item(t).rows.length-1
                        if tables.item(t).rows.item(r).cells.length >= 2
                            symb      = strtrim(tables.item(t).rows.item(r).cells.item(0).innerText);
                            label     = strtrim(tables.item(t).rows.item(r).cells.item(1).innerText);
                            % validate table content, first row must be {'Symbol','Name'}
                            if r == 0
                                if ~strcmp(symb,'Symbol') && ~strcmp(label,'Name')
                                    % table headers not valid, skip table
                                    break
                                end
                            else
                                % add line if line is valid (non empty label, symbol without whitespace)
                                if ~any(isspace(symb)) && ~isempty(label) && ~strcmpi(label,'n/a')
                                    symbols = [symbols;{symb,label}];
                                end
                            end
                        end
                    end
                end
            end
            % go to next page
            pagecount = pagecount + 1;
        end
        % add Index symbol to list
        if ~isempty(symbols)
            % remove duplicates
            [~,n] = unique(symbols(:,1));
            symbols = symbols(n,:);
            symbols = [{indexname,indexdesc};symbols];  
        end
        % close Internet Explorer
        h.Quit;
        
    catch ME
        % update status message
        set(handles.text_dataimport_datafeed_symbollookup_status,'String',ME.message)
        % enable 'fetch symbol' button
        set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
        return
    end
    
    % symbol table setup
    set(handles.uitable_dataimport_datafeed_symbols,'ColumnFormat',[{[]},'char','char']);
    pos = get(handles.uitable_dataimport_datafeed_symbols,'Position');
    if size(symbols,1) > 22
        set(handles.uitable_dataimport_datafeed_symbols,'ColumnWidth',{40,60,pos(3)-120});
    else
        set(handles.uitable_dataimport_datafeed_symbols,'ColumnWidth',{40,60,pos(3)-104});
    end
    set(handles.uitable_dataimport_datafeed_symbols,'ColumnEditable',[true,false,false]);
    set(handles.uitable_dataimport_datafeed_symbols,'ColumnName',{'','Symbol','Name'});
    
    % add data to table
    data = [repmat({true},size(symbols,1),1),symbols];
    set(handles.uitable_dataimport_datafeed_symbols,'Data',data);

    % update status
    if ~isempty(symbols)
        set(handles.text_dataimport_datafeed_symbollookup_status,'String','Finished!')
    else
        set(handles.text_dataimport_datafeed_symbollookup_status,'String','No symbols found!')
    end
    % enable 'fetch symbol' button
    set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
    % show download panel
    if ~isempty(symbols)
        set(handles.uipanel_dataimport_datafeed_download,'Title',[indexdesc,' Component Download']);
        set(handles.uipanel_dataimport_datafeed_download,'Visible','on');
        % Clear status fields
        set(handles.text_dataimport_datafeed_download_status,'String','');
        set(handles.text_dataimport_datafeed_download_errorstatus,'String','');
    end
    
end


% --- Executes on button press in button_dataimport_datafeed_downloadseries.
function button_dataimport_datafeed_downloadseries_Callback(hObject, eventdata, handles)
% hObject    handle to button_dataimport_datafeed_downloadseries (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Disable controls
    set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','off');
    disableImportDataPage(handles);
    
    % Clear download status
    set(handles.text_dataimport_datafeed_download_status,'String','');
    set(handles.text_dataimport_datafeed_download_errorstatus,'String','');
    
    % This callback is different, if download is already in progress
    % In this case, button text is 'Cancel Download', -> change to 'Cancelling..'
    button_text = get(handles.button_dataimport_datafeed_downloadseries,'String');
    if strcmp('Cancel Download',button_text)
        set(handles.button_dataimport_datafeed_downloadseries,'String','Cancelling..');
        drawnow expose;
        return
    end

    % Download selected prices series
    
    % get symbol list
    tabledata = get(handles.uitable_dataimport_datafeed_symbols,'Data');
    selection = cell2mat(tabledata(:,1));
    symbols = tabledata(:,2);
    priceslabels = tabledata(:,3);
    % only use selected
    symbols = symbols(selection);
    priceslabels = priceslabels(selection);
    
    % make sure there are at least three series selected
    if length(symbols) < 3
        % update status and return
        set(handles.text_dataimport_datafeed_download_status,'String','Please select at least 3 series from list');
        set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
        return
    end
    
    % get date range
    startdate = get(handles.edit_dataimport_datafeed_startdate,'String');
    enddate = get(handles.edit_dataimport_datafeed_enddate,'String');
    try   
        startdate = datenum(startdate,handles.datestringformat);
        enddate = datenum(enddate,handles.datestringformat);
    catch ME
        % update status and return
        set(handles.text_dataimport_datafeed_download_status,'String',['Start/End dates require format: ',handles.datestringformat]);
        set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
        return
    end
    if (enddate - startdate) < 7
        % update status and return
        set(handles.text_dataimport_datafeed_download_status,'String','Date range is less than one week');
        set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
        return
    end
        
    % Change button text to 'Cancel Download' and enable control
    set(handles.button_dataimport_datafeed_downloadseries,'String','Cancel Download');
    set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
    drawnow expose;
    
    % Fetch data with datafeed connection
    try
        % connect to Yahoo! Finance
        y = yahoo;
        i = 0;   % loop counter
        % Loop over all symbols, break if button text changes to 'Cancelling..'
        while (i < length(symbols)) && ...
              ~strcmp('Cancelling..',get(handles.button_dataimport_datafeed_downloadseries,'String'))
            % increment counter
            i = i + 1;  
            % update status
            set(handles.text_dataimport_datafeed_download_status,'String',['Downloading item ',num2str(i),'/',num2str(length(symbols))]);
            drawnow expose;
            % fetch single series
            try
                data = fetch(y,symbols{i},'Close',startdate,enddate);
            catch ME
                % add symbol name to error status text
                err_msg = get(handles.text_dataimport_datafeed_download_errorstatus,'String');
                err_msg{end+1,1} = ['Download of item ',priceslabels{i},' failed']; 
                set(handles.text_dataimport_datafeed_download_errorstatus,'String',err_msg);
                % do not add any data
                data = [];
            end
            if ~isempty(data)
                if i == 1
                    % create data matrix with dates in first column
                    rawseries = data;
                    % also collect labels
                    rawlabels = priceslabels(i);
                else
                    % find dates common in eprice and new data vector (and corresponding indeces)
                    [common_dates,ind_rawseries,ind_data] = intersect(rawseries(:,1),data(:,1));
                    % assign values
                    rawseries(ind_rawseries,end+1) = data(ind_data,2);
                    % save label
                    rawlabels{end+1,1} = priceslabels{i};
                end
            end
            % allow cancellation
            pause(0.05);
        end
    catch ME
        % update status/controls and return
        set(handles.text_dataimport_datafeed_download_status,'String',ME.message);
        set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
        set(handles.button_dataimport_datafeed_downloadseries,'String','Download Prices Series');
        set(handles.button_dataimport_datafeed_downloadseries,'Enable','on');
        return
    end
    
    % enable controls
    set(handles.button_dataimport_datafeed_fetchsymbols,'Enable','on');
    set(handles.button_dataimport_datafeed_downloadseries,'String','Download Prices Series');
    set(handles.button_dataimport_datafeed_downloadseries,'Enable','on');
    
    % verify if download has been cancelled by user
    if i < length(symbols)
        set(handles.text_dataimport_datafeed_download_status,'String','Download cancelled');
        return
    end
       
    % Update status
    set(handles.text_dataimport_datafeed_download_status,'String','Download finished!');
    
    % sort by date
    rawseries = sortrows(rawseries,1);

    % remove assets that do not have data for the full date range (remove if more than 5 values are empty)
    ind = sum(rawseries==0 | isnan(rawseries)) > 5;
    rawseries(:,ind) = [];
    rawlabels(ind)   = [];
    
    % remove all rows containing at least one either nan or 0 value
    [row,col] = find(rawseries==0 | isnan(rawseries));
    rawseries(row,:) = [];

    % separate items
    dates  = rawseries(:,1);
    if selection(1) == true
        index  = rawseries(:,2);
        indexlabel = rawlabels{1};
        prices = rawseries(:,3:end);
        priceslabels = rawlabels(2:end);
    else
        index = [];
        indexlabel = [];
        prices = rawseries(:,2:end);
        priceslabels = rawlabels;
    end
    
    % Store data as mat file
    save(['datafeed_import ',datestr(now,'yyyymmddHHMM')],'prices','index','dates','priceslabels','indexlabel');
    
    % Write data to table
    updateImportDataPage(handles,prices,index,dates,priceslabels,indexlabel);
    
end


% --- Write imported data to table
function updateImportDataPage(handles,prices,benchmark,dates,priceslabels,benchmarklabel)
% prices:          prices series
% benchmark:       benchmark series
% dates:           dates series
% priceslabels:    asset labels
% benchmarklabel:  benchmark label

    % add data to table
    if ~isempty(benchmark)
        series = [benchmark(:),prices];
        serieslabels = [benchmarklabel,priceslabels(:)'];
        set(handles.checkbox_importedseries_usefirstcolasbenchmark,'Value',1);  % valid benchmark as first column
    else
        series = prices;
        serieslabels = priceslabels;
        set(handles.checkbox_importedseries_usefirstcolasbenchmark,'Value',0);  % no benchmark
    end
    set(handles.uitable_importedseries_pricesseries,'ColumnName',serieslabels,'Data',series);
    if ~isempty(dates)
        datesstr = cellstr(datestr(dates,handles.datestringformat));
        set(handles.uitable_importedseries_pricesseries,'RowName',datesstr);
    else
        set(handles.uitable_importedseries_pricesseries,'RowName',[]);
    end

    % enable controls
    set(handles.uitable_importedseries_pricesseries,'Enable','on');
    set(handles.checkbox_importedseries_usefirstcolasbenchmark,'Enable','on');
    set(handles.button_importedseries_accept,'Enable','on');
    
end


function edit_dataimport_datafeed_startdate_Callback(hObject, eventdata, handles)
% hObject    handle to edit_dataimport_datafeed_startdate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_dataimport_datafeed_startdate as text
%        str2double(get(hObject,'String')) returns contents of edit_dataimport_datafeed_startdate as a double
end

% --- Executes during object creation, after setting all properties.
function edit_dataimport_datafeed_startdate_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_dataimport_datafeed_startdate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


function edit_dataimport_datafeed_enddate_Callback(hObject, eventdata, handles)
% hObject    handle to edit_dataimport_datafeed_enddate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_dataimport_datafeed_enddate as text
%        str2double(get(hObject,'String')) returns contents of edit_dataimport_datafeed_enddate as a double
end

% --- Executes during object creation, after setting all properties.
function edit_dataimport_datafeed_enddate_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_dataimport_datafeed_enddate (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end


% --- Executes on button press in button_dataimport_datafeed_download_selectall.
function button_dataimport_datafeed_download_selectall_Callback(hObject, eventdata, handles)
% hObject    handle to button_dataimport_datafeed_download_selectall (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % select all components in list
    tabledata = get(handles.uitable_dataimport_datafeed_symbols,'Data');
    if ~isempty(tabledata)
        tabledata(:,1) = {true};
        set(handles.uitable_dataimport_datafeed_symbols,'Data',tabledata);
    end
end

% --- Executes on button press in button_dataimport_datafeed_download_deselectall.
function button_dataimport_datafeed_download_deselectall_Callback(hObject, eventdata, handles)
% hObject    handle to button_dataimport_datafeed_download_deselectall (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % deselect all components in list
    tabledata = get(handles.uitable_dataimport_datafeed_symbols,'Data');
    if ~isempty(tabledata)
        tabledata(2:end,1) = {false};
        set(handles.uitable_dataimport_datafeed_symbols,'Data',tabledata);
    end
end


% --- Executes on button press in button_dataimport_excelfile_import.
function button_dataimport_excelfile_import_Callback(hObject, eventdata, handles)
% hObject    handle to button_dataimport_excelfile_import (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    %  Import data from Excel
    
    % Reset controls
    set(handles.text_dataimport_excelfile_status,'String','');    
    disableImportDataPage(handles);
    
    % get file name
    filename = get(handles.edit_dataimport_excelfile_filename,'String');
    if isempty(filename)
        % update status and return
        set(handles.text_dataimport_excelfile_status,'String','Please provide Excel filename');
        return
    end
    if ~exist(filename,'file')
        % update status and return
        set(handles.text_dataimport_excelfile_status,'String','Can''t find file specified');
        return
    end
    
    % get sheet name and import data from Excel
    sheetname = get(handles.edit_dataimport_excelfile_sheetname,'String');
    try
        if isempty(sheetname)
            [~,~,rawdata] = xlsread(filename);
        else
            [~,~,rawdata] = xlsread(filename,sheetname);
        end
    catch ME
        % update status and return
        set(handles.text_dataimport_excelfile_status,'String',ME.message);
        return
    end
    
    if isempty(rawdata)
        % update status and return
        set(handles.text_dataimport_excelfile_status,'String','Worksheet is empty');
        return
    end
    
    if (size(rawdata,1) < 10) || (size(rawdata,2) < 4)
        % update status and return
        set(handles.text_dataimport_excelfile_status,'String','Worksheet contains too few data');
        return
    end
    
    % get specified column headers 
    datesheadername = get(handles.edit_dataimport_excelfile_datesheadername,'String');
    
    % try to find dates column header name in first row
    ind = find(strcmp(rawdata(1,:),datesheadername));
    if ~isempty(ind) && isscalar(ind)
        dates = rawdata(2:end,ind);
        rawdata(:,ind) = [];
        % convert dates into serial dates if needed
        if ~isnumeric(dates{1})
            datestringformat = cellstr(get(handles.popupmenu_dataimport_excelfile_dateformat,'String'));
            sel = get(handles.popupmenu_dataimport_excelfile_dateformat,'Value');
            datestringformat = datestringformat{sel};
            try
                dates = datenum(dates,datestringformat);
            catch ME
                % update status and return
                set(handles.text_dataimport_excelfile_status,'String',ME.message);
                return
            end
            % save date string format persistently
            handles.datestringformat = datestringformat;
            guidata(hObject, handles);
        end
    else
        dates = [];
    end
    
    % show info wheter dates have been parsed
    if isempty(dates)
        msg = ['Note: No date header with name "',datesheadername,'" found'];
        set(handles.text_dataimport_excelfile_status,'String',msg);
    end
    
    % store prices series
    prices = rawdata(2:end,:);
    priceslabels = rawdata(1,:);

    % convert series to double arrays:
    % is first column string dates?
    bad_el = cellfun(@ischar,prices(:,1)) == true; % is it non-numeric?
    % if any left, are they NaN?
    for i = 1:size(prices,1)
        if (bad_el(i) == false) && (~isscalar(prices{i,1}) || isnan(prices{i,1}))
            bad_el(i) = true;
        end
    end
    if all(bad_el)
        % remove first column
        prices(:,1) = [];
        priceslabels(1) = [];
    end
    % remove bad rows (non-numeric items)
    [row,col] = find(cellfun(@isnumeric,prices) == false);
    prices(row,:) = [];  
    prices = cell2mat(prices);
    
    % Write data to table
    updateImportDataPage(handles,prices,[],dates,priceslabels,[]);
    
end


% --- Executes during object creation, after setting all properties.
function edit_dataimport_excelfile_datesheadername_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_dataimport_excelfile_datesheadername (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end



% --- Executes during object creation, after setting all properties.
function edit_dataimport_excelfile_sheetname_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_dataimport_excelfile_sheetname (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end



% --- Executes during object creation, after setting all properties.
function edit_dataimport_excelfile_filename_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_dataimport_excelfile_filename (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end

% --- Executes on button press in button_dataimport_excelfile_browse.
function button_dataimport_excelfile_browse_Callback(hObject, eventdata, handles)
% hObject    handle to button_dataimport_excelfile_browse (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Open file browser
%     [filename,pathname] = uigetfile({'*.xlsx;*.xls','Excel File'},'Please select file');
basePath=pwd;
pathChosen=[basePath,'\Finam\FinamData\'];
[filename,pathname] =uigetfile([pathChosen '*.xlsx;*.xls'],'Excel File');
    if ~isequal(filename,0)
        set(handles.edit_dataimport_excelfile_filename,'String',fullfile(pathname,filename));
    end
 
end



% --- Executes during object creation, after setting all properties.
function popupmenu_dataimport_excelfile_dateformat_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu_dataimport_excelfile_dateformat (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
end 


% --- Executes when entered data in editable cell(s) in uitable_dataseries_assetselection.
function uitable_dataseries_assetselection_CellEditCallback(hObject, eventdata, handles)
% hObject    handle to uitable_dataseries_assetselection (see GCBO)
% eventdata  structure with the following fields (see UITABLE)
%	Indices: row and column indices of the cell(s) edited
%	PreviousData: previous data for the cell(s) edited
%	EditData: string(s) entered by the user
%	NewData: EditData or its converted form set on the Data property. Empty if Data was not changed
%	Error: error string when failed to convert EditData to appropriate value for Data
% handles    structure with handles and user data (see GUIDATA)

    % Update portfolio component selection
    indices = eventdata.Indices;
    if indices(2) == 2
        % update asset list
        enableAsset(handles,indices(1),eventdata.NewData);
    end

end


% --- Select/deselect asset from list
function enableAsset(handles,asset,state)
% asset:  index of asset in list 
% state:  new state (true/false)

    % get table data
    data = get(handles.uitable_dataseries_assetselection,'Data');

    % update table if needed
    if data{asset,2} ~= state
        data{asset,2} = state;
        set(handles.uitable_dataseries_assetselection,'Data',data);
    end
        
    % do not allow to uncheck less than two assets
    sel = cell2mat(data(:,2));
    if (sum(sel) == 1) && (state == false)
        % reassign value
        data(asset,2) = {true};
        set(handles.uitable_dataseries_assetselection,'Data',data);
        return
    end
    % update asset selection in model
    handles.model.enableAsset(asset,state);
    % update interface
    updateDataSeriesPage(handles);
    updatePortfolioOptimizationPage(handles);

end


% --- Executes on button press in button_importedseries_accept.
function button_importedseries_accept_Callback(hObject, eventdata, handles)
% hObject    handle to button_importedseries_accept (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Get data from prices series table and update portfolio model
    data = get(handles.uitable_importedseries_pricesseries,'Data');
    if isempty(data)
        return
    end
    priceslabels = get(handles.uitable_importedseries_pricesseries,'ColumnName');
    if get(handles.checkbox_importedseries_usefirstcolasbenchmark,'Value')
        % use first column as benchmark index
        benchmark = data(:,1);
        prices = data(:,2:end);
        benchmarklabel = priceslabels{1};
        priceslabels(1) = [];
    else
        % no benchmark
        benchmark = [];
        prices = data;
        benchmarklabel = '';
    end
    dates = get(handles.uitable_importedseries_pricesseries,'RowName');
    if ~isempty(dates)
        dates = datenum(dates,handles.datestringformat);    
    end
    
    % Save data
    handles.model.importData(prices,benchmark,dates,priceslabels,benchmarklabel)
   
    % Update Data Series Pages
    clearDataSeriesPage(handles);
    updateDataSeriesPage(handles);
    
    % Update Portfolio Optimization Page
    clearPortfolioOptimizationPage(handles);
    updatePortfolioOptimizationPage(handles);
    
    % Clear Results Page
    clearResultsPage(handles);
    
    % Switch to Data Series page
    tab_handler(2,handles);


end


% --- Executes during object creation, after setting all properties.
function uipanel_dataimport_CreateFcn(hObject, eventdata, handles)
% hObject    handle to uipanel_dataimport (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
end


% --- Executes on button press in checkbox_importedseries_usefirstcolasbenchmark.
function checkbox_importedseries_usefirstcolasbenchmark_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox_importedseries_usefirstcolasbenchmark (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox_importedseries_usefirstcolasbenchmark
end


% --- Executes on button press in pushbutton24.
function pushbutton24_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton24 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%  Import data from Excel
  
    % get file name
    filename = get(handles.edit_constr_import_excelfile_filename,'String');
    if isempty(filename)
        % update status and return
        set(handles.edit_constr_import_excelfile_filename,'String','Please provide Excel filename');
        return
    end
    if ~exist(filename,'file')
        % update status and return
        set(handles.edit_constr_import_excelfile_filename,'String','Can''t find file specified');
        return
    end
%get data from xlsx file
    try
         [~,~,rawdata] = xlsread(filename);
    catch ME
        % update status and return
        set(handles.errorField1,'String',ME.message);
        return
    end
    
    if isempty(rawdata)
        % update status and return
        set(handles.errorField1,'String','Worksheet is empty');
        return
    end
    
    if (size(rawdata,1) < 1) || (size(rawdata,2) < 1)
        % update status and return
        set(handles.errorField1,'String','Worksheet contains too few data');
        return
    end
    
    % get specified column headers 
%     datesheadername = get(handles.edit_dataimport_excelfile_datesheadername,'String');
    
    % try to find dates column header name in first row
%     ind = find(strcmp(rawdata(1,:),datesheadername));

    %-----------------
    % Only run if data available
    if isempty(handles.model.getPrices)
        return
    end
    labels = handles.model.getPricesLabels;
    counter=1;
    for k=1:length(rawdata)
        [row,col]=find(strcmpi(labels,rawdata{k,1}));
        if ~isempty(row)
            % add an additional row to constraints table
            % if it is first row, all assets are enabled
            % otherwise all assets are disabled
            data = get(handles.uitable_portopt_genericconstraints,'Data');
            % get total number of assets
            numassets = length(handles.model.getAssetSelection);
            if isempty(data)
                data = [repmat({true},1,numassets),{'=',1,'Enabled'}];
            else
                
                if counter==1
                    data(2:end,:)=[];
                end
                data = [data;repmat({false},1,numassets),{'=',1,'Enabled'}]; 
                data{1,row}=false; %снимаем галку в поле, где указываются все ассеты
                data{counter+1,row}=true; %counter+1, т.к. заполняем со 2й строки
                data{counter+1,numassets+2}=rawdata{k,2}/100;
%                 str2num(rawdata{k,2})/100; % если эксель другой, то мб
%                 этот вариант нужен
            end
            data{1,numassets+2}=1-sum([data{2:end,numassets+2}]);
            set(handles.uitable_portopt_genericconstraints,'Data',data);
            counter=counter+1;
        end
    end
    %-----------------  

end



function edit_constr_import_excelfile_filename_Callback(hObject, eventdata, handles)
% hObject    handle to edit_constr_import_excelfile_filename (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_constr_import_excelfile_filename as text
%        str2double(get(hObject,'String')) returns contents of edit_constr_import_excelfile_filename as a double
end

% --- Executes during object creation, after setting all properties.
function edit_constr_import_excelfile_filename_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_constr_import_excelfile_filename (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
    if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
        set(hObject,'BackgroundColor','white');
    end
end

% --- Executes on button press in pushbutton25.
function pushbutton25_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton25 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
    % Open file browser
    [filename,pathname] = uigetfile({'*.xlsx;*.xls','Excel File'},'Please select file');
    if ~isequal(filename,0)
        set(handles.edit_constr_import_excelfile_filename,'String',fullfile(pathname,filename));
    end
end


% --- Executes on button press in portfolioWeights2file.
function portfolioWeights2file_Callback(hObject, eventdata, handles)
% hObject    handle to portfolioWeights2file (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

    % Options
    FileName = 'PortfolioWeights.xlsx';
    SheetName = 'Summary';
    Overwrite = true;
    
    
    % Delete existing report if necessary
    if Overwrite && exist(FileName,'file')
        delete(FileName);
    end
            % separate data from table into two columns
        data = get(handles.uitable_results_weights,'Data'); 
        for i = 1:size(data,1)
            ind1 = find(data{i,1}=='(',1,'last');
            ind2 = find(data{i,1}==')',1,'last');
            if ~isempty(ind1) && ~isempty(ind2)      
                temp2ndcolumn= data{i,1}(ind1+1:ind2-2);
                data{i,2}=temp2ndcolumn;
                data{i,1} = strtrim(data{i,1}(1:ind1-1));
            end
        end
    xlswrite(FileName,data);
 

end


% --- Executes when selected cell(s) is changed in uitable_dataseries_assetselection.
function uitable_dataseries_assetselection_CellSelectionCallback(hObject, eventdata, handles)
% hObject    handle to uitable_dataseries_assetselection (see GCBO)
% eventdata  structure with the following fields (see UITABLE)
%	Indices: row and column indices of the cell(s) currently selecteds
% handles    structure with handles and user data (see GUIDATA)
end
