function varargout = findCodeByIdsAndViceVerse(varargin)
% FINDCODEBYIDSANDVICEVERSE MATLAB code for findCodeByIdsAndViceVerse.fig
%      FINDCODEBYIDSANDVICEVERSE, by itself, creates a new FINDCODEBYIDSANDVICEVERSE or raises the existing
%      singleton*.
%
%      H = FINDCODEBYIDSANDVICEVERSE returns the handle to a new FINDCODEBYIDSANDVICEVERSE or the handle to
%      the existing singleton*.
%
%      FINDCODEBYIDSANDVICEVERSE('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in FINDCODEBYIDSANDVICEVERSE.M with the given input arguments.
%
%      FINDCODEBYIDSANDVICEVERSE('Property','Value',...) creates a new FINDCODEBYIDSANDVICEVERSE or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before findCodeByIdsAndViceVerse_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to findCodeByIdsAndViceVerse_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help findCodeByIdsAndViceVerse

% Last Modified by GUIDE v2.5 06-Mar-2016 22:33:55

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @findCodeByIdsAndViceVerse_OpeningFcn, ...
                   'gui_OutputFcn',  @findCodeByIdsAndViceVerse_OutputFcn, ...
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


% --- Executes just before findCodeByIdsAndViceVerse is made visible.
function findCodeByIdsAndViceVerse_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to findCodeByIdsAndViceVerse (see VARARGIN)

% Choose default command line output for findCodeByIdsAndViceVerse
handles.output = hObject;
%------------------------
allInfoIsExist = evalin('base', 'exist(''allInfo'')'); %get from WorkSpace if exist
% if (exist('allInfo','var')==0) %если нет, то скачаем
if (allInfoIsExist)
    allInfo = evalin('base', 'allInfo');
else
    allInfo=urlread('http://www.finam.ru/cache/icharts/icharts.js');
    assignin('base','allInfo',allInfo);
end
%----------------------------------
[startIndex1,endIndex1] = regexp(allInfo,'aEmitentIds = [.*var aEmitentNames = ');
aEmitentIds=allInfo(startIndex1+15:endIndex1-24);%actually this is not array, cause there as double as string elements
% aEmitentIds(length(aEmitentIds)-5000:end)

% a1=cellfun(@(s)sscanf(s,'%f,'), aEmitentIds);
%  a1=sscanf(aEmitentIds,'%f,');
tmp = regexp(aEmitentIds,'([^,]*)','tokens');
Ids = cat(2,tmp{:})';
Ids = strrep(Ids, char(39), '');%remove single quotes
assignin('base','Ids',Ids);
%----------------------------------
[startIndex2,endIndex2] = regexp(allInfo,'aEmitentNames = [.*var aEmitentCodes = ');
aEmitentNames  = allInfo(startIndex2+17:endIndex2-24);

% tmp = regexp(aEmitentNames,'([^,]*)','tokens');
Names = regexp(aEmitentNames, ''',''', 'split');
% Names = cat(2,tmp{:})';
Names=Names';
Names = strrep(Names, char(39), '');%remove single quotes
assignin('base','Names',Names);
%  aEmitentNames(length(aEmitentNames)-5000:end);
%  a2=sscanf(aEmitentNames,'%s,');

%----------------------------------
[startIndex3,endIndex3] = regexp(allInfo,'aEmitentCodes = [.*var aEmitentMarkets = ');
aEmitentCodes  = allInfo(startIndex3+17:endIndex3-26);

tmp = regexp(aEmitentCodes,'([^,]*)','tokens');
Codes = cat(2,tmp{:})';
Codes = strrep(Codes, char(39), '');%remove single quotes
assignin('base','Codes',Codes);
%  aEmitentCodes(length(aEmitentCodes)-5000:end);
%------------------------
[startIndex4,endIndex4] = regexp(allInfo,'aEmitentMarkets = [.*var aEmitentDecp = ');
aEmitentMarket  = allInfo(startIndex4+19:endIndex4-23);
tmp = regexp(aEmitentMarket,'([^,]*)','tokens');
Markets = cat(2,tmp{:})';
Markets = strrep(Markets, char(39), '');%remove single quotes
assignin('base','Markets',Markets);
%  aEmitentMarket(length(aEmitentMarket)-5000:end)
% Update handles structure
guidata(hObject, handles);

% UIWAIT makes findCodeByIdsAndViceVerse wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = findCodeByIdsAndViceVerse_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function code1_Callback(hObject, eventdata, handles)
% hObject    handle to code1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of code1 as text
%        str2double(get(hObject,'String')) returns contents of code1 as a double


% --- Executes during object creation, after setting all properties.
function code1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to code1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in buttonGetCodeById.
function buttonGetCodeById_Callback(hObject, eventdata, handles)
% hObject    handle to buttonGetCodeById (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
IdInput = get(handles.Id1,'string');
% c = num2str(str2double(a)+1);
CodesLocal = evalin('base', 'Codes'); %get from WorkSpace
IdsLocal = evalin('base', 'Ids'); %get from WorkSpace
MarketsLocal = evalin('base', 'Markets'); %get from WorkSpace
[row,col]=find(strcmpi(IdsLocal,IdInput)); %find neccessary element.
% CodesLocal(row)
% set(handles.textResultCodeById,'string',CodesLocal(row));
marketIndexOfStock=find(strcmp(MarketsLocal(row), '1'));%ищем только рынок акций (номер 1)
result=CodesLocal(row(marketIndexOfStock));
% result=CodesLocal(row);
if isempty(result)
   result='Nothing has found on stock market'; 
end
set(handles.textResultCodeById,'string',result);

function Id1_Callback(hObject, eventdata, handles)
% hObject    handle to Id1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of Id1 as text
%        str2double(get(hObject,'String')) returns contents of Id1 as a double


% --- Executes during object creation, after setting all properties.
function Id1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Id1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in buttonGetIdByCode.
function buttonGetIdByCode_Callback(hObject, eventdata, handles)
% hObject    handle to buttonGetIdByCode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
CodeInput = get(handles.code1,'string');
% c = num2str(str2double(a)+1);
CodesLocal = evalin('base', 'Codes'); %get from WorkSpace
IdsLocal = evalin('base', 'Ids'); %get from WorkSpace
MarketsLocal = evalin('base', 'Markets'); %get from WorkSpace
[row,col]=find(strcmpi(CodesLocal,CodeInput)); %find neccessary element.
% marketIndexOfStock=find(MarketsLocal(row)==1);
marketIndexOfStock=find(strcmp(MarketsLocal(row), '1'));%ищем только рынок акций (номер 1)
result=IdsLocal(row(marketIndexOfStock));
% result = cat(2,IdsLocal(row),CodesLocal(row))
if isempty(result)
   result='Nothing has found on stock market';
end
set(handles.textResultIdByCode,'string',result);



function textResultIdByCode_Callback(hObject, eventdata, handles)
% hObject    handle to textResultIdByCode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of textResultIdByCode as text
%        str2double(get(hObject,'String')) returns contents of textResultIdByCode as a double


% --- Executes during object creation, after setting all properties.
function textResultIdByCode_CreateFcn(hObject, eventdata, handles)
% hObject    handle to textResultIdByCode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function textResultCodeById_Callback(hObject, eventdata, handles)
% hObject    handle to textResultCodeById (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of textResultCodeById as text
%        str2double(get(hObject,'String')) returns contents of textResultCodeById as a double


% --- Executes during object creation, after setting all properties.
function textResultCodeById_CreateFcn(hObject, eventdata, handles)
% hObject    handle to textResultCodeById (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
