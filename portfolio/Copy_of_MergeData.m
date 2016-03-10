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
% resultArray=cell.empty;
namesArray=cell(1);
i=1;

for fileCount=1:length(dirContent)
    CurrentName=dirContent(fileCount).name;
    isCSVFile = regexp(CurrentName,'(.csv)'); %�������� ������ � csv
%     a=genvarname('x')
% eval([a ' =5'])
%     isTxtFile = regexp(CurrentName,'(.xls)'); % xls ����������. �������, ��� ������ �������� ������ � ����� ���
%     isSelfFile = regexp(CurrentName,'(testMerge.m)') ;
%     isIllegalFile = regexp(CurrentName,'(testMergeCommon.m)');
    % isSelfFile = regexp(CurrentName,'(deleteSomeSymbols.m)'); %���� �� �������
    if isempty(isCSVFile) %���� �� csv, ����������
        continue
    end
    fid = fopen([pathChosen,'\', CurrentName], 'rt');
    % data = textscan(fid,'%s %f %f %f %f %f %f');%, 'headerlines', 1, 'delimiter', ',');
    data = textscan(fid,'%s %s %s %s %s %s %s','headerlines', 1,  'delimiter', ';');
%     if fileCount==3
        new=data{1,1}; %��� ������ ������� �������� ������� � �����
        % �������� ������ �� �����, ��� ��� ������� ������ ����� �� �����
        % ���. (����� ������� 248 ����)
        if length(new)<80
            fclose(fid);
            continue
        end
        
        
        S = sprintf('%s*', new{:});
        N = sscanf(S, '%f*');
        dateArray=cell.empty;
        prepareDateArray=strtrim(cellstr(num2str(N)));%�������������� �� double � str
        for n=1:length(prepareDateArray)
           dateArray{n,1}=[prepareDateArray{n}(1:4),'/',prepareDateArray{n}(5:6),'/',prepareDateArray{n}(7:8)]; %��������� ��� ������ ������ ����          
        end   
        match1 = regexp(dateArray(1,1),'/01/');
        match2 = regexp(dateArray(end,1),'/12/');
        
        if  isempty(match1{1}) || isempty(match2{1}) 
            fclose(fid);
            CurrentName;
            continue
        end
        if length(dateArray)< 220
            length(dateArray);
        CurrentName;
        end
        resultDateArray{:,i}=dateArray;
        
        namesArray{i}= strrep(CurrentName,'.csv',''); %��� ��� ���������� .csv
%     end
    
    new=data{1,6}; %� 6� ������� ������� ���� ��������
    S = sprintf('%s*', new{:});
    N = sscanf(S, '%f*');
% convert double to text and then to cell
    tmp6 = textscan(sprintf('%i\n',N'),'%s');
    tmp6 = tmp6{1};
    resultArray{:,i}=tmp6;
    
    fclose(fid);
 i=i+1;
end
clear prepareDateArray;
pause(0.01);
set(handles.fieldProcess,'String','20%');
pause(0.01);
%������ ������� �������, ����� MICEX ��� �� ������ ����� (����� ��� �����
%�� �������)
j = find(strncmp(namesArray, 'MICEX', 3)); %
if j ~=1
    k=1;%������ ������� (� ������� ������)
    %������ � ������� ���
    resultDateArray=resultDateArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
    namesArray=namesArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
    %������ � ������� ������
    resultArray=resultArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
%     namesArray=namesArray(:,[1:k-1,j,k+1:j-1,k,j+1:end]);
end

[fin, arrayOfIndexesAfterIntersect{1:length(resultDateArray)}] = mintersect(resultDateArray{:});

resultDataArray=double.empty;
for i=1:length(resultDateArray)
    %�������� ������� ������ � ����� ������� ����� � �������, �.�.
    %������ ������������� ����� ���� �����.
   [nonCommonVal,nonCommonInd]=setdiff(resultDateArray{i},fin);
   deleteIndexes=zeros(1);
   for k=1:length(nonCommonInd)
      if nonCommonInd(k) < arrayOfIndexesAfterIntersect{i}(1) %|| nonCommonInd(k) > arrayOfIndexesAfterIntersect{i}(end)
         deleteIndexes=[deleteIndexes k]; 
      end
   end
   deleteIndexes=deleteIndexes(2:end); % ������� ������ ����.  
   %������� �� ��������� ������� ������ � ��������� ������� ��������
   resultArray{i}(nonCommonInd(deleteIndexes))=[];
   resultDateArray{i}(nonCommonInd(deleteIndexes))=[];
end


%����������� ������ �� ���������� ��������, ������� ����� ���������� ��
%���� ���. ��� ���� ��� ��-������ ����� �������, �.�. ���� ����� ��������,
%��� ����� = 12.
lastDateInArray=zeros(1,length(resultDateArray));
for k=1:length(resultDateArray)
  lastDateTemp=char(resultDateArray{k}(end)); 
  [startIndex1,endIndex1] = regexp(lastDateTemp,'/12/');
  day=lastDateTemp(endIndex1+1:end);
  lastDateInArray(1,k)= str2num(day);
end
[valOfMin,indOfMin]=min(lastDateInArray);
benchmarkLastElement=resultDateArray{indOfMin}(end);

resultDataArray=double.empty;
for i=1:length(resultDateArray)
if i==51
    a=234;
end
%���� �������������� ������� � �������?
    [row,col]=find(strcmpi(resultDateArray{i},benchmarkLastElement));
    if ~isempty(row) %����
        %�� ���������?
        currentLastElement=resultDateArray{i}(end);
        if currentLastElement{:}==benchmarkLastElement{:} %������ ����� ����� ��� ������� ������
%         do nothing. All is ok.
        else
            resultArray{i}=resultArray{i}(1:row); %���������, �������� ���������
            resultDateArray{i}=resultDateArray{i}(1:row);
        end
    else %������ �������� � ������� ���
        for j=1:31
           searchForDay=valOfMin-j;
         
           lastDateTemp=char(resultDateArray{i}(end)); %����� ��� ������� ���� '2012/12/30/
           [startIndex1,endIndex1] = regexp(lastDateTemp,'/12/');
           lastDateTemp(endIndex1+1:end)=num2str(searchForDay); %��������������� ������� ����
           [row1,col1]=find(strcmpi(resultDateArray{i},lastDateTemp));
            %����������������� ���� ���� � �������?
           if ~isempty(row1) %����
               %�������� ������� �����������
               for k=1:(valOfMin-searchForDay) %��� �� =j
                   
                   %cell � double
                   i1=row1+1; % ������ �� �������� ����� ��������
                   op1=resultArray{i}(i1-1,1);
                   S = sprintf('%s*', op1{:});
                   op1 = sscanf(S, '%f*');
                   
                   op2=resultArray{i}(i1,1);
                   S = sprintf('%s*', op2{:});
                   op2 = sscanf(S, '%f*');
                   
                   
                   b=( op1+op2 )/2;%�������� ��� �������
                   %��������� ��������
                   b=num2str(b);
                   resultArray{i} = [resultArray{i}(1:i1-1,1)',b,resultArray{i}(i1:end,1)']';
                   
                   %��������� ����.   
                   [startIndex2,endIndex2] = regexp(lastDateTemp,'/12/');
                   lastDateTemp(endIndex1+1:end)=num2str(searchForDay+k);
                   i1=row1+1; % ������ �� �������� ����� ��������
                   b=lastDateTemp;
                   resultDateArray{i}=[resultDateArray{i}(1:i1-1,1)',b,resultDateArray{i}(i1:end,1)']';
                   
               end
               resultArray{i}=resultArray{i}(1:i1); %��������� ������ ������, �������� ���������
               resultDateArray{i}=resultDateArray{i}(1:i1);%��������� ������ ���, �������� ���������
               break
           else % ���
               %���������� ������

                
           end
        end
    end
end

% ������� ������������ ����� � ������ ����� �������� �� ������� ���.
[val,ind]=max(cellfun('length',resultDateArray));
pause(0.01);
set(handles.fieldProcess,'String','40%');
pause(0.01);
for i=1:length(resultDateArray)
%    resultDateArray{ind} - ����� ������� ������
%    resultDateArray{i} - ������� ������
   [nonCommonVal,nonCommonInd]=setdiff(resultDateArray{ind},resultDateArray{i});
  
   %���������� cell � double

   resultArrayIntersect{:,i}=  resultArray{i};
   tmp1=resultArrayIntersect{:,i};

   S = sprintf('%s**', tmp1{:});
   N = sscanf(S, '%f**');
   if i==65
       a=230948;
   end;
   for k=1:length(nonCommonInd)
%        resultDataArray(:,i) = N; % ��� ������� ��� ���������� ������
       i1=nonCommonInd(k); % ������ �� �������� ����� ��������
       b=( N(i1-1,1)+N(i1,1) )/2;
       N = [N(1:i1-1,1)',b,N(i1:end,1)']';
       
   end
%    length(N);
   resultDataArray(:,i) = N;
end
pause(0.01);
set(handles.fieldProcess,'String','70%');
pause(0.01);
%  resultArray{1}(arrayOfIndexesAfterIntersect{1}) - ��������
% ����� � xlsx
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
    
% ����� �������� ���������� � ���������
xlRange = 'B2';
% xlswrite(filename,resultDataArray,sheet,xlRange);
xlswrite(pathForCreate,resultDataArray,sheet,xlRange);
pause(0.01);
set(handles.fieldProcess,'String','80%');
pause(0.01);
% ����� ������ ��������
xlRange = 'A1';
namesArray=['Date' namesArray];
xlswrite(pathForCreate,namesArray,sheet,xlRange);
pause(0.01);
set(handles.fieldProcess,'String','90%');
pause(0.01);
% ����� ����
xlRange = 'A2';
xlswrite(pathForCreate,resultDateArray{ind},sheet,xlRange);
set(handles.fieldProcess,'String','100%. ���� �����������');

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
  
  
