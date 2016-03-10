clear all
clc

basePath=pwd;
path1='all2011';
path2='1.0.2012-1.0.2013';
path3='1.1.2014-1.1.2015';
pathChosen=[basePath,'\..\Finam\FinamData\data\', path3];
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
        if length(new)<89
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

% [fin, arrayOfIndexesAfterIntersect{1:length(resultDateArray)}] = mintersect(resultDateArray{:});
% 
% resultDataArray=double.empty;
% for i=1:length(resultDateArray)
%     %�������� ������� ������ � ����� ������� ����� � �������, �.�.
%     %������ ������������� ����� ���� �����.
%    [nonCommonVal,nonCommonInd]=setdiff(resultDateArray{i},fin);
%    deleteIndexes=zeros(1);
%    for k=1:length(nonCommonInd)
%       if nonCommonInd(k) < arrayOfIndexesAfterIntersect{i}(1) %|| nonCommonInd(k) > arrayOfIndexesAfterIntersect{i}(end)
%          deleteIndexes=[deleteIndexes k]; 
%       end
%    end
%    deleteIndexes=deleteIndexes(2:end); % ������� ������ ����.  
%    %������� �� ��������� ������� ������ � ��������� ������� ��������
%    resultArray{i}(nonCommonInd(deleteIndexes))=[];
%    resultDateArray{i}(nonCommonInd(deleteIndexes))=[];
% end


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


for i=1:length(resultDateArray)
if i==92
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
        for j=1:90
           searchForDay=valOfMin-j;
%            if searchForDay<10
%               searchForDay=['0' num2str(searchForDay)]; 
%            end
%            lastDateTemp=char(resultDateArray{i}(end)); %����� ��� ������� ���� '2012/12/30/
%            [startIndex1,endIndex1] = regexp(lastDateTemp,'/12/');
%            lastDateTemp(endIndex1+1:end)=num2str(searchForDay); %��������������� ������� ����
%            
%            
           formatIn = 'yyyy/mm/dd';
           benchAsDate=datenum(benchmarkLastElement{:},formatIn);
           benchAsDateAddedDay=addtodate(benchAsDate, -j, 'day');
           
           formatOut = 'yyyy/mm/dd';
           benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut);

           [row1,col1]=find(strcmpi(resultDateArray{i},benchAsDateAddedDayString));
            %����������������� ���� ���� � �������?
           if ~isempty(row1) %����
               %�������� ������� �����������
               for k=1:(j) %��� �� =j
                   
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
%                    [startIndex2,endIndex2] = regexp(lastDateTemp,'/12/');
                   
                   
                   formatIn = 'yyyy/mm/dd';
                   benchAsDate=datenum(resultDateArray{i}(row1),formatIn);
                   benchAsDateAddedDay=addtodate(benchAsDate, k, 'day');

                   formatOut = 'yyyy/mm/dd';
                   benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut); 
                   
%                    lastDateTemp(endIndex1+1:end)=num2str(searchForDay+k);
                   i1=row1+1; % ������ �� �������� ����� ��������
                   b=benchAsDateAddedDayString;
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
%������ ����������� ������ �� ������� ��������, ������� ����� ���������� ��
%���� ���. ��� ���� ��� ��-������ ����� ������, �.�. ���� ����� ��������,
%��� ����� = 1.
firstDateInArray=zeros(1,length(resultDateArray));
for k=1:length(resultDateArray)
  firstDateTemp=char(resultDateArray{k}(1)); 
  [startIndex1,endIndex1] = regexp(firstDateTemp,'/01/');
  day=firstDateTemp(endIndex1+1:end);
  firstDateInArray(1,k)= str2num(day);
end
[valOfMin,indOfMin]=max(firstDateInArray);
benchmarkFirstElement=resultDateArray{indOfMin}(1);


for i=1:length(resultDateArray)
if i==6
    a=234;
end
%���� �������������� ������� � �������?
    [row,col]=find(strcmpi(resultDateArray{i},benchmarkFirstElement));
    if ~isempty(row) %����
        %�� ������?
        currentFirstElement=resultDateArray{i}(1);
        if currentFirstElement{:}==benchmarkFirstElement{:} %{}������ ����� ����� ��� ������� ������
%         do nothing. All is ok.
        else
            resultArray{i}=resultArray{i}(row:end); %���������, �������� ���������
            resultDateArray{i}=resultDateArray{i}(row:end);
        end
    else %������ �������� � ������� ���
        for j=1:90
           searchForDay=valOfMin-j;
         
%            firstDateTemp=char(resultDateArray{i}(1)); %����� ��� ������� ���� '2012/12/30/
%            [startIndex1,endIndex1] = regexp(firstDateTemp,'/01/');
%            firstDateTemp(endIndex1+1:end)=num2str(searchForDay); %��������������� ������� ����
%            [row1,col1]=find(strcmpi(resultDateArray{i},firstDateTemp));
           formatIn = 'yyyy/mm/dd';
           benchAsDate=datenum(benchmarkFirstElement{:},formatIn);
           benchAsDateAddedDay=addtodate(benchAsDate, +j, 'day');
           
           formatOut = 'yyyy/mm/dd';
           benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut);

           [row1,col1]=find(strcmpi(resultDateArray{i},benchAsDateAddedDayString));           
%����������������� ���� ���� � �������?
           if ~isempty(row1) %����
               %�������� ������� �����������
               for k=1:(j) %��� �� =j
                   
                   %cell � double
                   i1=row1; % ������ �� �������� ����� ��������
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
%                    [startIndex2,endIndex2] = regexp(lastDateTemp,'/12/');
                   
                   
                   formatIn = 'yyyy/mm/dd';
                   benchAsDate=datenum(resultDateArray{i}(row1),formatIn);
                   benchAsDateAddedDay=addtodate(benchAsDate, -1, 'day');

                   formatOut = 'yyyy/mm/dd';
                   benchAsDateAddedDayString=datestr(benchAsDateAddedDay,formatOut); 
                   
%                    lastDateTemp(endIndex1+1:end)=num2str(searchForDay+k);
                   i1=row1; % ������ �� �������� ����� ��������
                   b=benchAsDateAddedDayString;
                   resultDateArray{i}=[resultDateArray{i}(1:i1-1,1)',b,resultDateArray{i}(i1:end,1)']';
                   
               end
               resultArray{i}=resultArray{i}(i1:end); %��������� ������ ������, �������� ���������
               resultDateArray{i}=resultDateArray{i}(i1:end);%��������� ������ ���, �������� ���������
               break
           else % ���
               %���������� ������

                
           end
        end
    end
end
% ������� ������������ ����� � ������ ����� �������� �� ������� ���.
[val,ind]=max(cellfun('length',resultDateArray));
completeVecDate=m_union(resultDateArray{:});
for i=1:length(resultDateArray)
%    resultDateArray{ind} - ����� ������� ������
%    resultDateArray{i} - ������� ������
   [nonCommonVal,nonCommonInd]=setdiff(completeVecDate,resultDateArray{i});
  
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

%  resultArray{1}(arrayOfIndexesAfterIntersect{1}) - ��������
% ����� � xlsx
[pathstr,name,ext] = fileparts(pathChosen);
filename = [name ext,'.xlsx'];
sheet = 1;
pathForCreate=[basePath,'\..\Finam\FinamData\', filename];
% pathChosen=[basePath,'\..\Finam\FinamData\data\', path3];
Overwrite = true;   
% Delete existing report if necessary
if Overwrite && exist(filename,'file')
   delete(filename);
end
    
% ����� �������� ���������� � ���������
xlRange = 'B2';
% xlswrite(filename,resultDataArray,sheet,xlRange);
xlswrite(pathForCreate,resultDataArray,sheet,xlRange);
% ����� ������ ��������
xlRange = 'A1';
namesArray=['Date' namesArray];
xlswrite(pathForCreate,namesArray,sheet,xlRange);
% ����� ����
xlRange = 'A2';
xlswrite(pathForCreate,completeVecDate,sheet,xlRange);



