function varargout = MergeData(varargin)
% MERGEDATA MATLAB code for MergeData.fig
%      MERGEDATA, by itself, creates a new MERGEDATA or raises the existing
%      singleton*.
%
%      H = MERGEDATA returns the handle to a new MERGEDATA or the handle to
%      the existing singleton*.
%
%      MERGEDATA('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in MERGEDATA.M with the given input arguments.
%
%      MERGEDATA('Property','Value',...) creates a new MERGEDATA or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before MergeData_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to MergeData_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help MergeData

% Last Modified by GUIDE v2.5 09-Mar-2016 15:48:02

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @MergeData_OpeningFcn, ...
                   'gui_OutputFcn',  @MergeData_OutputFcn, ...
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


% --- Executes just before MergeData is made visible.
function MergeData_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to MergeData (see VARARGIN)

% Choose default command line output for MergeData
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes MergeData wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = MergeData_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in buttonDone.
function buttonDone_Callback(hObject, eventdata, handles)
% hObject    handle to buttonDone (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
set(handles.fieldProcess,'String','10%');
pause(0.01);
folderPathChosen = get(handles.FieldFolderName,'String');

pathChosen=folderPathChosen;
dirContent = dir(pathChosen);

[pathstr,name,ext] = fileparts(pathChosen);
foldernameForParse=[name ext];
[stMonth1,stMonth2]=regexp(foldernameForParse,'^.{2}\.');
startMonth=foldernameForParse(stMonth2+1:stMonth2+2); %выделили начальный месяц
startMonth=str2double(startMonth);
startYear=foldernameForParse(stMonth2+4:stMonth2+7);
startYear=str2double(startYear);

[endMonth1,endMonth2]=regexp(foldernameForParse,'\.[0-9]{4}$');
endMonth=foldernameForParse(endMonth1-2:endMonth1-1); %выделили конечный месяц
endMonth=str2double(endMonth);
endYear=foldernameForParse(endMonth1+1:end);
endYear=str2double(endYear);
% resultArray=cell.empty;
namesArray=cell(1);
i=1;
%флаги начального и конечного месяца. Был ли хотя бы где-то переход на
%extra месяц. Если был, то смещаем границу на него.
% flagStart=0;
% flagEnd=0;
for fileCount=1:length(dirContent)
    CurrentName=dirContent(fileCount).name;
    isCSVFile = regexp(CurrentName,'(.csv)'); %работаем только с csv
%     a=genvarname('x')
% eval([a ' =5'])
%     isTxtFile = regexp(CurrentName,'(.xls)'); % xls пропускаем. Считаем, что больше ненужных файлов в папке нет
%     isSelfFile = regexp(CurrentName,'(testMerge.m)') ;
%     isIllegalFile = regexp(CurrentName,'(testMergeCommon.m)');
    % isSelfFile = regexp(CurrentName,'(deleteSomeSymbols.m)'); %себя не удаляем
    if isempty(isCSVFile) %если не csv, пропускаем
        continue
    end
    fid = fopen([pathChosen,'\', CurrentName], 'rt');
    % data = textscan(fid,'%s %f %f %f %f %f %f');%, 'headerlines', 1, 'delimiter', ',');
    data = textscan(fid,'%s %s %s %s %s %s %s','headerlines', 1,  'delimiter', ';');
%     if fileCount==3
        new=data{1,1}; %при первом прогоне забиваем столбец с датой
        % выбираем только те файлы, где есть данные
        % (всего порядка 248 торговых дней)
        if length(new)<30
            fclose(fid);
            continue
        end
        
        
        S = sprintf('%s*', new{:});
        N = sscanf(S, '%f*');
        dateArray=cell.empty;
        prepareDateArray=strtrim(cellstr(num2str(N)));%преобразование из double в str
        for n=1:length(prepareDateArray)
           dateArray{n,1}=[prepareDateArray{n}(1:4),'/',prepareDateArray{n}(5:6),'/',prepareDateArray{n}(7:8)]; %подгоняем под нужный формат даты          
        end   
        match1 = regexp(dateArray(1,1),sprintf('%.4d/%.2d/',startYear, startMonth));
        match2 = regexp(dateArray(end,1),sprintf('%.4d/%.2d/',endYear, endMonth));
        %ищем на месяц назад, если в этом месяце ничего нет. Так, например,
        %если задано 1 января, то записи будут только на декабрь.
%         endMonthExtra=endMonth-1;
%         if endMonthExtra==0
%             endMonthExtra=12;
%         end
%         match2extra = regexp(dateArray(end,1),sprintf('/%.2d/', endMonthExtra));
         %ищем на месяц вперед, если в этом месяце ничего нет. Так, например,
        %если задано 31 января, то записи будут скорее всего только на февраль.
%         startMonthExtra=startMonth+1;
%         if startMonthExtra==13
%             startMonthExtra=1;
%         end
% %         match1extra = regexp(dateArray(1,1),sprintf('/%.2d/', startMonthExtra));
        %Если в обоих первых и последних месяцах не будет найдено
        %указанного месяца, то пропустить файл
%         if  ( isempty(match1{1}) && isempty(match1extra{1}) )  || (isempty(match2{1}) && isempty(match2extra{1}) )
        if  isempty(match1{1}) || isempty(match2{1})
                    
            fclose(fid);
            CurrentName;
            continue
        end
        if isempty(match2{1})
            flagEnd=1;        
        end
        if isempty(match1{1})
            flagStart=1;        
        end
        resultDateArray{:,i}=dateArray;
        
        namesArray{i}= strrep(CurrentName,'.csv',''); %имя без расширения .csv
%     end
    
    new=data{1,6}; %в 6м столбце искомые цены закрытия
    S = sprintf('%s*', new{:});
    N = sscanf(S, '%f*');
% convert double to text and then to cell
    tmp6 = textscan(sprintf('%i\n',N'),'%s');
    tmp6 = tmp6{1};
    resultArray{:,i}=tmp6;
    
    fclose(fid);
 i=i+1;
end
% if flagEnd==1
%     endMonth=endMonthExtra; 
% end
% if flagStart==1
%     startMonth=startMonthExtra; 
% end
clear prepareDateArray;
pause(0.01);
set(handles.fieldProcess,'String','20%');
pause(0.01);
%меняем местами столбцы, чтобы MICEX был на первом месте (чтобы его взять
%за образец)
j = find(strncmp(namesArray, 'MICEX', 3)); %
if j ~=1
    k=1;%первый столбец (с которым меняем)
    %меняем в массиве дат
    resultDateArray=resultDateArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
    namesArray=namesArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
    %меняем в массиве данных
    resultArray=resultArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
%     namesArray=namesArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
end

%укорачиваем массив по последнему значению, которое имеет наименьшую из
%всех дат. при этом это по-любому будет декабрь, т.к. выше стоит проверка,
%что месяц = 12.
lastDateInArray=zeros(1,length(resultDateArray));
for k=1:length(resultDateArray)
 
    %Проверяем, был ли переход на extra месяц. Если был, то смещаем границу
  [indExt1,indExt2]=regexp(resultDateArray{k},sprintf('/%.2d/', endMonth));
  IndexForShift=max(find(~cellfun(@isempty,indExt1)));
  lastDateTemp=char(resultDateArray{k}(IndexForShift));
  
  [startIndex1,endIndex1] = regexp(lastDateTemp,sprintf('/%.2d/', endMonth));
  day=lastDateTemp(endIndex1+1:end);
%   k
  lastDateInArray(1,k)= str2num(day);
  
  resultDateArray{k}=resultDateArray{k}(1:IndexForShift);
  resultArray{k}=resultArray{k}(1:IndexForShift);
end
[valOfMin,indOfMin]=min(lastDateInArray);
benchmarkLastElement=resultDateArray{indOfMin}(end);


for i=1:length(resultDateArray)
if i==6
    a=234;
end
%Есть бенчмарковский элемент в массиве?
    [row,col]=find(strcmpi(resultDateArray{i},benchmarkLastElement));
    if ~isempty(row) %есть
        %Он последний?
        currentLastElement=resultDateArray{i}(end);
        if currentLastElement{:}==benchmarkLastElement{:} %скобки чтобы взять сам элемент внутри
%         do nothing. All is ok.
        else
            resultArray{i}=resultArray{i}(1:row); %укоротили, отбросив последние
            resultDateArray{i}=resultDateArray{i}(1:row);
        end
    else %такого элемента в массиве нет
        for j=1:180
           
           formatIn = 'yyyy/mm/dd';
           benchAsDate=datenum(benchmarkLastElement{:},formatIn);
           benchAsDateAddedDay=addtodate(benchAsDate, -j, 'day');
           
           formatOut = 'yyyy/mm/dd';
           benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut);

           [row1,col1]=find(strcmpi(resultDateArray{i},benchAsDateAddedDayString));
            %сконструированная дата есть в массиве?
           if ~isempty(row1) %есть
               %заменяем пробелы полусуммами
               for k=1:(j) %оно же =j
                   
                   %cell в double
                   i1=row1+1; % индекс по которому нужно вставить
                   op1=resultArray{i}(i1-1,1);
                   S = sprintf('%s*', op1{:});
                   op1 = sscanf(S, '%f*');
                   
                   op2=resultArray{i}(i1,1);
                   S = sprintf('%s*', op2{:});
                   op2 = sscanf(S, '%f*');
                   
                   
                   b=( op1+op2 )/2;%значение для вставки
                   %вставляем значение
                   b=num2str(b);
                   resultArray{i} = [resultArray{i}(1:i1-1,1)',b,resultArray{i}(i1:end,1)']';
                   
                   %вставляем дату.   
%                    [startIndex2,endIndex2] = regexp(lastDateTemp,'/12/');
                   
                   
                   formatIn = 'yyyy/mm/dd';
                   benchAsDate=datenum(resultDateArray{i}(row1),formatIn);
                   benchAsDateAddedDay=addtodate(benchAsDate, k, 'day');

                   formatOut = 'yyyy/mm/dd';
                   benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut); 
                   
                   i1=row1+1; % индекс по которому нужно вставить
                   b=benchAsDateAddedDayString;
                   resultDateArray{i}=[resultDateArray{i}(1:i1-1,1)',b,resultDateArray{i}(i1:end,1)']';
                   
               end
               resultArray{i}=resultArray{i}(1:i1); %укоротили массив данных, отбросив последние
               resultDateArray{i}=resultDateArray{i}(1:i1);%укоротили массив дат, отбросив последние
               break
           else % нет
               %продолжаем искать

                
           end
        end
    end
end
%теперь укорачиваем массив по первому значению, которое имеет наибольшую из
%всех дат. при этом это по-любому будет январь, т.к. выше стоит проверка,
%что месяц = 1.
firstDateInArray=zeros(1,length(resultDateArray));
for k=1:length(resultDateArray)
%    if k==113
%       a=5; 
%    end
  %Проверяем, был ли переход на extra месяц. Если был, то смещаем границу
  [indExt1,indExt2]=regexp(resultDateArray{k},sprintf('/%.2d/', startMonth));
  IndexForShift=min(find(~cellfun(@isempty,indExt1)));
  firstDateTemp=char(resultDateArray{k}(IndexForShift));
  
  [startIndex1,endIndex1] = regexp(firstDateTemp,sprintf('/%.2d/', startMonth));
  day=firstDateTemp(endIndex1+1:end);
  firstDateInArray(1,k)= str2num(day);
  
  resultDateArray{k}=resultDateArray{k}(IndexForShift:end);
  resultArray{k}=resultArray{k}(IndexForShift:end);
end
[valOfMin,indOfMin]=max(firstDateInArray);
benchmarkFirstElement=resultDateArray{indOfMin}(1);


for i=1:length(resultDateArray)
if i==6
    a=234;
end
%Есть бенчмарковский элемент в массиве?
    [row,col]=find(strcmpi(resultDateArray{i},benchmarkFirstElement));
    if ~isempty(row) %есть
        %Он первый?
        currentFirstElement=resultDateArray{i}(1);
        if currentFirstElement{:}==benchmarkFirstElement{:} %{}скобки чтобы взять сам элемент внутри
%         do nothing. All is ok.
        else
            resultArray{i}=resultArray{i}(row:end); %укоротили, отбросив последние
            resultDateArray{i}=resultDateArray{i}(row:end);
        end
    else %такого элемента в массиве нет
        for j=1:180

           formatIn = 'yyyy/mm/dd';
           benchAsDate=datenum(benchmarkFirstElement{:},formatIn);
           benchAsDateAddedDay=addtodate(benchAsDate, +j, 'day');
           
           formatOut = 'yyyy/mm/dd';
           benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut);

           [row1,col1]=find(strcmpi(resultDateArray{i},benchAsDateAddedDayString));           
%сконструированная дата есть в массиве?
           if ~isempty(row1) %есть
               %заменяем пробелы полусуммами
               for k=1:(j) %оно же =j
                   
                   %cell в double
                   i1=row1; % индекс по которому нужно вставить
                   op1=resultArray{i}(i1-1,1);
                   S = sprintf('%s*', op1{:});
                   op1 = sscanf(S, '%f*');
                   
                   op2=resultArray{i}(i1,1);
                   S = sprintf('%s*', op2{:});
                   op2 = sscanf(S, '%f*');
                   
                   
                   b=( op1+op2 )/2;%значение для вставки
                   %вставляем значение
                   b=num2str(b);
                   resultArray{i} = [resultArray{i}(1:i1-1,1)',b,resultArray{i}(i1:end,1)']';
                   
                   %вставляем дату.   
%                    [startIndex2,endIndex2] = regexp(lastDateTemp,'/12/');
                   
                   
                   formatIn = 'yyyy/mm/dd';
                   benchAsDate=datenum(resultDateArray{i}(row1),formatIn);
                   benchAsDateAddedDay=addtodate(benchAsDate, -1, 'day');

                   formatOut = 'yyyy/mm/dd';
                   benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut); 
                   
                   i1=row1; % индекс по которому нужно вставить
                   b=benchAsDateAddedDayString;
                   resultDateArray{i}=[resultDateArray{i}(1:i1-1,1)',b,resultDateArray{i}(i1:end,1)']';
                   
               end
               resultArray{i}=resultArray{i}(i1:end); %укоротили массив данных, отбросив последние
               resultDateArray{i}=resultDateArray{i}(i1:end);%укоротили массив дат, отбросив последние
               break
           else % нет
               %продолжаем искать

                
           end
        end
    end
end

% a=cell.empty; for k=1:length(resultDateArray)
%  a=[a;resultDateArray{k}(1)];
%  end
% for k=1:length(resultDateArray)
%  [row,col]=find(strcmpi(resultDateArray{k},'2013/12/18'));
%  if ~isempty(row)
%      k
%      row
%      col
%  end
%  end
%'2013/12/18'
% [row,col]=find(strcmpi([resultDateArray{end}],'2013/12/18'))

% находим максимальную длину и индекс этого элемента из массива дат.
[val,ind]=max(cellfun('length',resultDateArray));
completeVecDate=m_union(resultDateArray{:});
pause(0.01);
set(handles.fieldProcess,'String','40%');
pause(0.01);
for i=1:length(resultDateArray)
%    resultDateArray{ind} - самый длинный вектор
%    resultDateArray{i} - текущий вектор
   [nonCommonVal,nonCommonInd]=setdiff(completeVecDate,resultDateArray{i});
  
   %Превращаем cell в double

   resultArrayIntersect{:,i}=  resultArray{i};
   tmp1=resultArrayIntersect{:,i};

   S = sprintf('%s**', tmp1{:});
   N = sscanf(S, '%f**');
   if i==24
       a=230948;
   end;
   for k=1:length(nonCommonInd)
%        resultDataArray(:,i) = N; % уже готовый для расширения массив
       i1=nonCommonInd(k); % индекс по которому нужно вставить
       i;
       k;
          if k==89
       a=230948;
   end;
       b=( N(i1-1,1)+N(i1,1) )/2;
       N = [N(1:i1-1,1)',b,N(i1:end,1)']';
       
   end
%    length(N);
   resultDataArray(:,i) = N;
end
pause(0.01);
set(handles.fieldProcess,'String','70%');
pause(0.01);
%  resultArray{1}(arrayOfIndexesAfterIntersect{1}) - работает
% Вывод в xlsx
[pathstr,name,ext] = fileparts(pathChosen);
filename = [name ext,'.xlsx'];
sheet = 1;
basePath=pwd;
pathForCreate=[basePath,'\Finam\FinamData\', filename];
Overwrite = true;   
% Delete existing report if necessary
if Overwrite && exist(filename,'file')
   delete(filename);
end
    
% Вывод основной информации о стоимости
xlRange = 'B2';
% xlswrite(filename,resultDataArray,sheet,xlRange);
xlswrite(pathForCreate,resultDataArray,sheet,xlRange);
pause(0.01);
set(handles.fieldProcess,'String','80%');
pause(0.01);
% Вывод строки названий
xlRange = 'A1';
namesArray=['Date' namesArray];
xlswrite(pathForCreate,namesArray,sheet,xlRange);
pause(0.01);
set(handles.fieldProcess,'String','90%');
pause(0.01);
% Вывод даты
xlRange = 'A2';
xlswrite(pathForCreate,completeVecDate,sheet,xlRange);
set(handles.fieldProcess,'String','100%. Файл сформирован');




function FieldFolderName_Callback(hObject, eventdata, handles)
% hObject    handle to FieldFolderName (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of FieldFolderName as text
%        str2double(get(hObject,'String')) returns contents of FieldFolderName as a double


% --- Executes during object creation, after setting all properties.
function FieldFolderName_CreateFcn(hObject, eventdata, handles)
% hObject    handle to FieldFolderName (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in buttonBrowse.
function buttonBrowse_Callback(hObject, eventdata, handles)
% hObject    handle to buttonBrowse (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
basePath=pwd;
pathChosen=[basePath,'\Finam\FinamData\data\'];
pathName =uigetdir(pathChosen);

   if ~isequal(pathName,0)
        set(handles.FieldFolderName,'String',fullfile(pathName));
    end
  
  
