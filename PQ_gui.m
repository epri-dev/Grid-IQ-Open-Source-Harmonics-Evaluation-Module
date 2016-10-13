function varargout = PQ_gui(varargin)

% Copyright © 2012-2016, Electric Power Research Institute, Inc.
% All rights reserved.
% Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
% 1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
% 2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
% 3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
% THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE

% Last Modified by GUIDE v2.5 29-Jul-2016 14:28:53


if nargin == 0  % LAUNCH GUI
    
    fig = openfig(mfilename,'reuse');
 % Use system color scheme for figure:
    set(fig,'Color',get(0,'DefaultUicontrolBackgroundColor'));
    scrzx=get(0,'ScreenSize');
    
    set(fig, 'OuterPosition', [0 scrzx(4)-600 800 600]);
%     set(fig, 'OuterPosition', [0 .05 1 .95]);
       
    % Generate a structure of handles to pass to callbacks, and store it.
    handles = guihandles(fig);

    logoimage = imread('EPRILogo2.jpg');
    image(logoimage,'parent',handles.logoaxes);
    set(handles.logoaxes,'Visible','off','DataAspectRatio',[1 1 1]);

    
    %These are default variables that are used in the GUI.  When changed, they
    %do not repopulate in the GUI window
    handles.mydir = cd;
    handles.ckt = '';
    handles.mydir = [cd,'\',handles.ckt];
    handles.ScanBus = '';
    handles.DistortionBus = '';
    handles.MonitoredElement = '';
    handles.ScanElement='None';
    handles.DistortionMonitoredElement = '';
    handles.ScanType = 'Positive';
    handles.FSphase = '1';
    handles.loadmult = 1;
    handles.SeriesRL = 0;
    handles.AllCapConfigs=0;
    handles.AllCapConfigs2=0;
    handles.AddedLoadDistortionData=zeros(10,100,2);
    handles.Freq = 60;
    handles.Harm = 3;

    set(handles.SplashScreen,'Visible','on')
    
    
    [DSSStartOK, DSSObj, DSSText] = DSSStartup(cd);
    handles.DSSObj = DSSObj;
    handles.DSSText = DSSText;
    DSSText.command = ['set defaultbasefrequency=',num2str(handles.Freq)];
    
    set(handles.ControlPanel,'Visible','off')
    set(handles.uipanel4,'Visible','off')
    set(handles.uipanel11,'Visible','off')
    set(handles.uipanel13,'Visible','off')
    set(handles.AboutPanel,'Visible','off')
%     set(handles.UseNLtext,'Visible','off')

    guidata(fig, handles)
    if nargout > 0
        varargout{1} = fig;
    end
    
elseif ischar(varargin{1}) % INVOKE NAMED SUBFUNCTION OR CALLBACK
    
    try
        if (nargout)
            [varargout{1:nargout}] = feval(varargin{:}); % FEVAL switchyard
        else
            feval(varargin{:}); % FEVAL switchyard
        end
    catch
        disp(lasterr);
    end
    
end


%| ABOUT CALLBACKS:
%| GUIDE automatically appends subfunction prototypes to this file, and
%| sets objects' callback properties to call them through the FEVAL
%| switchyard above. This comment describes that mechanism.
%|
%| Each callback subfunction declaration has the following form:
%| <SUBFUNCTION_NAME>(H, EVENTDATA, HANDLES, VARARGIN)
%|
%| The subfunction name is composed using the object's Tag and the
%| callback type separated by '_', e.g. 'slider2_Callback',
%| 'figure1_CloseRequestFcn', 'axis1_ButtondownFcn'.
%|
%| H is the callback object's handle (obtained using GCBO).
%|
%| EVENTDATA is empty, but reserved for future use.
%|
%| HANDLES is a structure containing handles of components in GUI using
%| tags as fieldnames, e.g. handles.figure1, handles.slider2. This
%| structure is created at GUI startup using GUIHANDLES and stored in
%| the figure's application data using GUIDATA. A copy of the structure
%| is passed to each callback.  You can store additional information in
%| this structure at GUI startup, and you can change the structure
%| during callbacks.  Call guidata(h, handles) after changing your
%| copy to replace the stored original so that subsequent callbacks see
%| the updates. Type "help guihandles" and "help guidata" for more
%| information.
%|
%| VARARGIN contains any extra arguments you have passed to the
%| callback. Specify the extra arguments by editing the callback
%| property in the inspector. By default, GUIDE sets the property to:
%| <MFILENAME>('<SUBFUNCTION_NAME>', gcbo, [], guidata(gcbo))
%| Add any extra arguments after the last argument, before the final
%| closing parenthesis.

% --- Executes on button press in CktSummary.
% function CktSummary_Callback(hObject, eventdata, handles)
% var1 = importdata([handles.mydir,'\summary.csv'],',',1);
% f = figure(40);
% set(f,'Name','Circuit Summary','Position',[200 200 1000 200]);
% cnames = {' MaxPuVoltage','MinPuVoltage','TotalMW','TotalMvar','kWLosses','pctLosses','kvarLosses','Frequency'};
% uitable('Parent',f,'Data',var1.data(5:12),'ColumnName',cnames,...
%     'RowName',([]),'Position',[20 20 960 160]);

% --- Executes on button press in CompileCircuit.
function CompileCircuit_Callback(hObject, eventdata, handles)
if get(handles.CompileCircuit,'Value')==1
    CapBusList = [];
    CapBusListMod = [];
    count=0;
    captabledata=get(handles.CapData,'Data');
    MaxXcoord=0;
    MinXcoord=0;
    MaxYcoord=0;
    MinYcoord=0;
    for ii=1:length(handles.ElementNames)
        handles.DSSObj.ActiveCircuit.Buses(char(strtok(handles.DSSObj.ActiveCircuit.CktElements(char(handles.ElementNames(ii))).bus(1),'.')));
        count = count + handles.DSSObj.ActiveCircuit.ActiveBus.Coorddefined;
        if handles.DSSObj.ActiveCircuit.ActiveBus.Coorddefined
            if count==1
                MaxXcoord=handles.DSSObj.ActiveCircuit.ActiveBus.x;
                MinXcoord=handles.DSSObj.ActiveCircuit.ActiveBus.x;
                MaxYcoord=handles.DSSObj.ActiveCircuit.ActiveBus.y;
                MinYcoord=handles.DSSObj.ActiveCircuit.ActiveBus.y;
            end
            MaxXcoord=max(MaxXcoord,handles.DSSObj.ActiveCircuit.ActiveBus.x);
            MinXcoord=min(MinXcoord,handles.DSSObj.ActiveCircuit.ActiveBus.x);
            MaxYcoord=max(MaxYcoord,handles.DSSObj.ActiveCircuit.ActiveBus.y);
            MinYcoord=min(MinYcoord,handles.DSSObj.ActiveCircuit.ActiveBus.y);
        end
    end
    if count>1
        for ii=1:length(handles.CapNames)
            %         handles.DSSObj.ActiveCircuit.Buses(char(strtok(handles.DSSObj.ActiveCircuit.CktElements(char(handles.CapNames(ii))).bus(1),'.')));
            %         if handles.DSSObj.ActiveCircuit.ActiveBus.Coorddefined==0;
            %             errordlg({['No Bus Coordinates Defined for Capacitor in row ',num2str(ii)]},'Capacitor Buscoords Not Defined','modal')
            %         end
            %         CapBus = strtok(handles.DSSObj.ActiveCircuit.CktElements(char(handles.CapNames(ii))).bus(1),'.');
            CapBus = char(captabledata(ii,2));
            CapBus = strtok(CapBus,'.');
            CapBusListMod{ii}=char(CapBus);
            handles.DSSObj.ActiveCircuit.SetActiveBus(CapBus);
            CapBusXcoord(ii)=handles.DSSObj.ActiveCircuit.ActiveBus.x;
            CapBusYcoord(ii)=handles.DSSObj.ActiveCircuit.ActiveBus.y;
            if length(find(strcmp(char(CapBus),CapBusListMod)))==1
                for ii=1:ii
                    CapBusList = [CapBusList,' ',char(CapBus)];
                end
            end
        end
        handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
        handles.DSSText.command = ['set loadmult=',num2str(handles.loadmult)];
        handles.DSSText.command = 'Solve mode=snap';
        %Daisy Cap Display
        Daisysize=1;
%         if (MaxXcoord-MinXcoord)/(MaxYcoord-MinYcoord)>100||(MaxYcoord-MinYcoord)/(MaxXcoord-MinXcoord)>100
%             Daisysize=Daisysize^3
%         end
        handles.DSSText.command = ['Set daisysize=',num2str(Daisysize)];
        handles.DSSText.command = ['Plot daisy power 1phLinestyle=3 max=',num2str(handles.MaxPowerPlot),' dots=N subs=yes buslist=[',char(CapBusList),']'];
        %----
        %DSS Marker Cap Display
        %     AddedCapData=get(handles.AddedCapacitors,'Data');
        %     if cell2mat(AddedCapData(1,1))~=0
        %         for ii=1:size(AddedCapData,1)
        %             if cell2mat(AddedCapData(ii,3))
        %                 handles.DSSText.Command = ['New Capacitor.',char(AddedCapData(ii,1)),' bus=',char(AddedCapData(ii,1)),' kvar=1'];
        %             end
        %         end
        %     end
        %     handles.DSSText.command = 'Solve mode=snap';
        %     handles.DSSText.command = ['Set MarkCapacitors=yes'];
        %     handles.DSSText.command = ['Plot circuit 1phLinestyle=3 max=5000 dots=N subs=yes'];
        %----
        %Matlab Feeder plot with caps identified
        %     var1 = importdata([handles.mydir,'\DSSGraph_Output.CSV'],',',2);
        %     figure(100)
        %     for z=1:size(var1.data,1)
        %         line = [var1.data(z,1) var1.data(z,2);var1.data(z,3) var1.data(z,4)];
        %         plot(line(:,1),line(:,2),'k','linewidth',var1.data(z,6)), hold on;
        %     end
        %     scatter(CapBusXcoord,CapBusYcoord,'filled')
        %     text(CapBusXcoord,CapBusYcoord,{char(CapBusListMod)});
        % %     axis off;
        %     hold off
        %----
        guidata(hObject,handles)
    else
        errordlg({'No Bus Coordinates Defined for this Circuit.';'No Plot Available.'},'Buscoords Not Defined','modal')
    end
end

% --- Executes when entered data in editable cell(s) in CapData.
function CapData_CellEditCallback(hObject, eventdata, handles)
if eventdata.Indices(2)==7
    row = eventdata.Indices(1);
    captabledata=get(handles.CapData,'Data');
    if strcmp(eventdata.NewData,'wye')
        captabledata{row,9}='Solidly Grounded';
    elseif strcmp(eventdata.NewData,'delta')
        captabledata{row,9}='Ungrounded';
    elseif strcmp(eventdata.NewData,'LN')
        captabledata{row,9}='Solidly Grounded';
    elseif strcmp(eventdata.NewData,'LL')
        captabledata{row,9}='Ungrounded';
    end
    handles.CapDataTable=captabledata;
    set(handles.CapData,'Data',handles.CapDataTable);
end
if eventdata.Indices(2)==8
    row = eventdata.Indices(1);
    captabledata=get(handles.CapData,'Data');
    CapPhase=eventdata.NewData;
    if isempty(strfind(CapPhase,'.1.2.3'))==0||isempty(strfind(CapPhase,'.1.2'))==0||isempty(strfind(CapPhase,'.2.3'))==0||isempty(strfind(CapPhase,'.1.3'))==0
        CapkV=captabledata{row,10}*sqrt(3);
    else
        CapkV=captabledata{row,10};
    end
    captabledata{row,11}=CapkV;
    handles.CapDataTable=captabledata;
    set(handles.CapData,'Data',handles.CapDataTable);
end
guidata(hObject,handles)

% --- Executes when entered data in editable cell(s) in AddedCapacitors.
function AddedCapacitors_CellEditCallback(hObject, eventdata, handles)
tabledata=get(handles.AddedCapacitors,'Data');
if eventdata.Indices(2)==4
    if eventdata.NewData==1
        fid = fopen([handles.mydir,'\DSSViewIntercom.txt'],'r');
        user_entry = fgetl(fid);
        fclose(fid);
        NewCapBus = char(handles.DSSObj.ActiveCircuit.CktElements(char(user_entry)).bus(1));
        row = eventdata.Indices(1);
        %Clear previous cap from list if one has already been added.
        if cell2mat(tabledata(row,3))==1
            CapNames=handles.CapNames;
            captabledata=get(handles.CapData,'Data');
            ind=find(strcmp(captabledata(:,1),['Capacitor.',char(tabledata(row,1))]));
            captabledata(ind,:)=[];
            CapNames(ind,:)=[];
            handles.CapNames=CapNames;
            handles.CapDataTable=captabledata;
            set(handles.CapData,'Data',handles.CapDataTable);
            set(handles.CompileCircuit,'Value',0)
        end
        tabledata(row,1)=cellstr(NewCapBus);
        tabledata(row,3)=[{false}];
        tabledata(row,4)=[{false}];
        set(handles.AddedCapacitors,'Data',tabledata)
    end
end
if eventdata.Indices(2)==3
    if eventdata.NewData==1
        row = eventdata.Indices(1);
        CapNames=handles.CapNames;
        handles.CapNames = [CapNames;{['Capacitor.',char(tabledata(row,1))]}];
        captabledata=get(handles.CapData,'Data');
        [token,remain] = strtok(char(tabledata(row,1)),'.');
        BuskV = handles.DSSObj.ActiveCircuit.Buses(cell2mat(strtok(tabledata(row,1),'.'))).kV;
        if isempty(remain)
            CapPhase= '.1.2.3';
        else
            CapPhase = remain;
        end
        if isempty(strfind(CapPhase,'.1.2.3'))==0||isempty(strfind(CapPhase,'.1.2'))==0||isempty(strfind(CapPhase,'.2.3'))==0||isempty(strfind(CapPhase,'.1.3'))==0
            CapkV=BuskV*sqrt(3);
        else
            CapkV=BuskV;
        end
        handles.CapDataTable=[captabledata;['Capacitor.',char(tabledata(row,1))],tabledata(row,1),true,300,false,0,'wye',CapPhase,' ',BuskV,CapkV];
        set(handles.CapData,'Data',handles.CapDataTable);
        set(handles.CompileCircuit,'Value',0)
    end
    if eventdata.NewData==0
        row = eventdata.Indices(1);
        CapNames=handles.CapNames;
        captabledata=get(handles.CapData,'Data');
        ind=find(strcmp(captabledata(:,1),['Capacitor.',char(tabledata(row,1))]));
        captabledata(ind,:)=[];
        CapNames(ind,:)=[];
        handles.CapNames=CapNames;
        handles.CapDataTable=captabledata;
        set(handles.CapData,'Data',handles.CapDataTable);
        set(handles.CompileCircuit,'Value',0)
    end
end
guidata(hObject,handles)

function loadmult_Callback(hObject, eventdata, handles)
user_entry = str2double(get(hObject,'string'));
handles.loadmult = user_entry;
set(handles.CompileCircuit,'Value',0)

handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
handles.DSSText.command = ['set loadmult=',num2str(handles.loadmult)];
handles.DSSText.command = 'Solve mode=snap';
delete([handles.mydir,'\*ummary.csv']) 
handles.DSSText.command = 'Export Summary Summary.csv';
var1 = importdata([handles.mydir,'\summary.csv'],',',1);
handles.MaxPowerPlot = var1.data(7)*1000;
% f = figure(40);
% set(f,'Name','Circuit Summary','Position',[200 200 1000 200]);
% cnames = {' MaxPuVoltage','MinPuVoltage','TotalMW','TotalMvar','kWLosses','pctLosses','kvarLosses','Frequency'};
% uitable('Parent',f,'Data',var1.data(5:12),'ColumnName',cnames,...
%     'RowName',([]),'Position',[20 20 960 160]);
set(handles.text62,'String',[num2str(round(var1.data(7)*10)/10),' MW, ',num2str(round(var1.data(8)*10)/10),' Mvar']);

guidata(hObject,handles)

function Freq_Callback(hObject, eventdata, handles)
user_entry = str2double(get(hObject,'string'));
handles.Freq = user_entry;
handles.DSSText.command = ['set defaultbasefrequency=',num2str(handles.Freq)];
set(handles.CompileCircuit,'Value',0)

handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
handles.DSSText.command = ['set loadmult=',num2str(handles.loadmult)];
handles.DSSText.command = 'Solve mode=snap';
delete([handles.mydir,'\*ummary.csv']) 
handles.DSSText.command = 'Export Summary Summary.csv';
var1 = importdata([handles.mydir,'\summary.csv'],',',1);
handles.MaxPowerPlot = var1.data(7)*1000;
% f = figure(40);
% set(f,'Name','Circuit Summary','Position',[200 200 1000 200]);
% cnames = {' MaxPuVoltage','MinPuVoltage','TotalMW','TotalMvar','kWLosses','pctLosses','kvarLosses','Frequency'};
% uitable('Parent',f,'Data',var1.data(5:12),'ColumnName',cnames,...
%     'RowName',([]),'Position',[20 20 960 160]);
set(handles.text62,'String',[num2str(round(var1.data(7)*10)/10),' MW, ',num2str(round(var1.data(8)*10)/10),' Mvar']);

guidata(hObject,handles)

function Isc_Callback(hObject, eventdata, handles)
user_entry = str2double(get(hObject,'string'));
handles.Isc = user_entry;
set(handles.CompileCircuit,'Value',0)

handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
handles.DSSText.command = ['edit vsource.* isc3=',num2str(handles.Isc)];
handles.DSSText.command = ['edit vsource.* isc1=',num2str(handles.Isc1)];
handles.DSSText.command = 'Solve mode=snap';
delete([handles.mydir,'\*ummary.csv']) 
handles.DSSText.command = 'Export Summary Summary.csv';
var1 = importdata([handles.mydir,'\summary.csv'],',',1);
handles.MaxPowerPlot = var1.data(7)*1000;
set(handles.text62,'String',[num2str(round(var1.data(7)*10)/10),' MW, ',num2str(round(var1.data(8)*10)/10),' Mvar']);
guidata(hObject,handles)

function Isc1_Callback(hObject, eventdata, handles)
user_entry = str2double(get(hObject,'string'));
handles.Isc1 = user_entry;
set(handles.CompileCircuit,'Value',0)

handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
handles.DSSText.command = ['edit vsource.* isc3=',num2str(handles.Isc)];
handles.DSSText.command = ['edit vsource.* isc1=',num2str(handles.Isc1)];
handles.DSSText.command = 'Solve mode=snap';
delete([handles.mydir,'\*ummary.csv']) 
handles.DSSText.command = 'Export Summary Summary.csv';
var1 = importdata([handles.mydir,'\summary.csv'],',',1);
handles.MaxPowerPlot = var1.data(7)*1000;
set(handles.text62,'String',[num2str(round(var1.data(7)*10)/10),' MW, ',num2str(round(var1.data(8)*10)/10),' Mvar']);
guidata(hObject,handles)

function SeriesRL_Callback(hObject, eventdata, handles)
user_entry = str2double(get(hObject,'string'));
if user_entry>100, user_entry=100; end
if user_entry<0, user_entry=0; end
handles.SeriesRL = abs(100-user_entry);
guidata(hObject,handles)


function MonitoredElement_Callback(hObject, eventdata, handles)
fid = fopen([handles.mydir,'\DSSViewIntercom.txt'],'r');
TLINE = fgetl(fid);
fclose(fid);
if strcmp(handles.ScanElement,'None')
    handles.ScanElement = TLINE;
else
    handles.ScanElement = [handles.ScanElement,',',TLINE];
end
handles.MonitoredElement=textscan(handles.ScanElement,'%s', 'delimiter', ',');
% NewScanBus = char(handles.DSSObj.ActiveCircuit.CktElements(char(handles.MonitoredElement{1}{1})).bus(1));
% handles.ScanBus = NewScanBus;
% hand = guihandles;
% set(hand.ScanBus,'String',NewScanBus)
set(handles.EditMonElement,'String',[char(handles.ScanElement)]);
guidata(hObject,handles)


function EditMonElement_Callback(hObject, eventdata, handles)
handles.ScanElement = get(hObject,'string');
handles.MonitoredElement = textscan(handles.ScanElement,'%s', 'delimiter', ',');
% NewScanBus = char(handles.DSSObj.ActiveCircuit.CktElements(char(handles.MonitoredElement{1}{1})).bus(1));
% handles.ScanBus = NewScanBus;
% hand = guihandles;
% set(hand.ScanBus,'String',NewScanBus)
guidata(hObject,handles)

% --- Executes on button press in ClearResponseElement.
function ClearResponseElement_Callback(hObject, eventdata, handles)
hand = guihandles;
set(hand.EditMonElement,'String','')
handles.MonitoredElement = '';
handles.ScanElement='None';
guidata(hObject,handles)


function DistortionMonitoredElement_Callback(hObject, eventdata, handles)
fid = fopen([handles.mydir,'\DSSViewIntercom.txt'],'r');
user_entry = fgetl(fid);
fclose(fid);
handles.DistortionMonitoredElement = user_entry;
NewScanBus = char(handles.DSSObj.ActiveCircuit.CktElements(char(handles.DistortionMonitoredElement)).bus(2));
handles.DistortionBus = NewScanBus;
set(handles.edit56,'String',[char(handles.DistortionMonitoredElement)]);
set(handles.edit57,'String',[char(handles.DistortionBus)]);
guidata(hObject,handles)

function edit56_Callback(hObject, eventdata, handles)
user_entry = get(hObject,'string');
handles.DistortionMonitoredElement = user_entry;

foundelement = 0;
for ii=1:length(handles.ElementNames)
    [tok,rem] = strtok(handles.ElementNames(ii),'.');
    if strcmp(tok,'Line') || strcmp(tok,'Transformer') || strcmp(tok,'Reactor')
        if strcmp(handles.ElementNames(ii),char(user_entry))
            foundelement = 1;
            NewScanBus = char(handles.DSSObj.ActiveCircuit.CktElements(char(handles.DistortionMonitoredElement)).bus(2));
            handles.DistortionBus = NewScanBus;
            set(handles.edit57,'String',[char(handles.DistortionBus)])
        end
    end
end
%Lookup in the added elements table.
if foundelement == 0
    tabledata = get(handles.AddedElements,'Data');
    for ii=1:size(tabledata,1)
        if strcmpi(['Transformer.',char(tabledata(ii,1))],char(user_entry))
            foundelement = 1;
            NewScanBus = char(tabledata(ii,12));
            handles.DistortionBus = NewScanBus;
            set(handles.edit57,'String',[char(handles.DistortionBus)])
        end
    end
end
if foundelement == 0
    msgbox({'The Monitored Element may not exist.';'If this is a new transformer added in 1c),';' the POI bus must be entered first.'},'Warning','modal')
end

guidata(hObject,handles)

function edit57_Callback(hObject, eventdata, handles)
user_entry = get(hObject,'string');
handles.DistortionBus = user_entry;

foundelement = 0;
for ii=1:length(handles.ElementNames)
    [tok,rem] = strtok(handles.ElementNames(ii),'.');
    if strcmp(tok,'Line') || strcmp(tok,'Transformer') || strcmp(tok,'Reactor')
        elementbus2 = char(strtok(handles.DSSObj.ActiveCircuit.CktElements(char(handles.ElementNames(ii))).bus(2),'.'));
        if strcmp(elementbus2,char(user_entry))
            handles.DistortionMonitoredElement = char(handles.ElementNames(ii));
            set(handles.edit56,'String',[char(handles.DistortionMonitoredElement)]);
            foundelement = 1;
        end
    end
end
if foundelement == 0
    tabledata = get(handles.AddedElements,'Data');
    for ii=1:size(tabledata,1)
        if strcmpi(char(tabledata(ii,12)),char(user_entry))
            handles.DistortionMonitoredElement = ['Transformer.',char(tabledata(ii,1))];
            set(handles.edit56,'String',['Transformer.',char(tabledata(ii,1))]);
            foundelement = 1;
        end
    end
end
if foundelement == 0
    handles.DistortionMonitoredElement = '';
    set(handles.edit56,'String',[char(handles.DistortionMonitoredElement)]);
    msgbox({'Please provide the Monitored Element that terminates with this Bus.';'If this Bus is part of a new transformer added in (1c),';'manually enter the Monitored Element which is Transformer.XX';'where XX is the Filename of the Source Type'},'Warning','modal')
end
guidata(hObject,handles)


% --- Executes on button press in LoadScanBus.
function LoadScanBus_Callback(hObject, eventdata, handles)
fid = fopen([handles.mydir,'\DSSViewIntercom.txt'],'r');
user_entry = fgetl(fid);
fclose(fid);
NewScanBus = char(handles.DSSObj.ActiveCircuit.CktElements(char(user_entry)).bus(1));
handles.ScanBus = NewScanBus;
hand = guihandles;
set(hand.ScanBus,'String',NewScanBus)
guidata(hObject,handles)


function ScanBus_Callback(hObject, eventdata, handles)
user_entry = get(hObject,'string');
handles.ScanBus = user_entry;
guidata(hObject,handles)

% --- Executes on selection change in Circuit.
function Circuit_Callback(hObject, eventdata, handles)
tabledata = {''    []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  [];
    ''        []    []    ''    ''    ''    [false]    [false]    [false]    [false]   [false] ''  []  ''  []  []  []  []};
set(handles.AddedElements,'Data',tabledata)
tabledata = {''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false];
    ''    ''    [false]    [false]};
set(handles.AddedCapacitors,'Data',tabledata)

user_entry = get(hObject,'string');
number = get(hObject,'Value');
if strcmp(char(user_entry{number}),'Load User Defined')
    handles.mydir = uigetdir(cd,'Select Folder Containing Feeder Files');
    handles.ckt = handles.mydir(length(cd)+2:length(handles.mydir));
elseif strcmp(char(user_entry{number}),' ')
    msgbox({'No Citcuit Selected'},'Warning','modal')
    return
else
    handles.ckt = user_entry{number};
    handles.mydir = [cd,'\',handles.ckt];
end

handles.date = date;
mkdir([handles.mydir,'/Figs/',handles.date])
mkdir([handles.mydir,'/AddedElements/'])
handles.DSSText.command = ['set defaultbasefrequency=',num2str(60)];
handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
handles.DSSText.command = 'Solve mode=snap';
Freq = handles.DSSObj.ActiveCircuit.Solution.Frequency;
handles.Freq = Freq;
handles.DSSText.command = ['set defaultbasefrequency=',num2str(Freq)];
hand = guihandles;
set(hand.Freq,'String',Freq)
set(hand.DisplayCkt,'String',handles.ckt)
set(hand.ScanBus,'String','')
handles.DistortionBus = '';
set(hand.EditMonElement,'String','')
handles.MonitoredElement = '';
set(hand.edit56,'String','')
set(hand.edit57,'String','')
handles.ScanElement='None';
delete([handles.mydir,'\*ummary.csv']) 
handles.DSSText.command = 'Export Summary Summary.csv';
var1 = importdata([handles.mydir,'\Summary.csv'],',',1);
handles.MaxPowerPlot = var1.data(7)*1000;
set(handles.text62,'String',[num2str(round(var1.data(7)*10)/10),' MW, ',num2str(round(var1.data(8)*10)/10),' Mvar']);

ElementNames = handles.DSSObj.ActiveCircuit.AllElementNames;
ElementNames_mod = strtok(ElementNames,'.');
capind = find(strcmp(ElementNames_mod,'Capacitor'));
CapNames = ElementNames(capind);
handles.CapNames = CapNames;

for ii=1:length(handles.CapNames)
    CapBusvar(ii,1) = handles.DSSObj.ActiveCircuit.CktElements(char(handles.CapNames(ii))).bus(1);
    CapStatus{ii,1} = handles.DSSObj.ActiveCircuit.CktElements(char(handles.CapNames(ii))).enabled;
end

iCap = handles.DSSObj.ActiveCircuit.Capacitors.First;
ii=1;
while iCap > 0
    handles.DSSObj.ActiveCircuit.Capacitors.name;
    FilterStatus{ii,1} = false;
    CapHarm{ii,1} = 0;
    CapConn{ii,1} = handles.DSSObj.ActiveCircuit.ActiveDSSElement.Properties('conn').Val;
    CapPhaseVar = handles.DSSObj.ActiveCircuit.ActiveDSSElement.Properties('phases').Val;
    CapKvar{ii,1} = handles.DSSObj.ActiveCircuit.Capacitors.kvar;
    CapkV{ii,1} = handles.DSSObj.ActiveCircuit.Capacitors.kV;
    BuskV{ii,1} = handles.DSSObj.ActiveCircuit.Buses(cell2mat(strtok(CapBusvar(ii,1),'.'))).kV;
    [token, remain] = strtok(handles.DSSObj.ActiveCircuit.ActiveDSSElement.Properties('bus1').Val,'.');
    if isempty(remain)==1
        CapPhase{ii,1} = '.1.2.3';
    else
        CapPhase{ii,1} = remain;
    end
    if isempty(strfind(CapPhase{ii,1},'.1.2.3'))==0
        CapPhasevar=3;
    elseif isempty(strfind(CapPhase{ii,1},'.1.2'))==0||isempty(strfind(CapPhase{ii,1},'.2.3'))==0||isempty(strfind(CapPhase{ii,1},'.1.3'))==0
        CapPhasevar=2;
    else
        CapPhasevar=1;
    end
    if CapKvar{ii,1}<1
        Cap_cuf = strtok(handles.DSSObj.ActiveCircuit.ActiveDSSElement.Properties('cuf').Val,'(');
        Cap_cuf = str2double(strtok(Cap_cuf,')'))/1000000;
        CapkV{ii,1} = BuskV{ii,1};
        if CapPhasevar==3||CapPhasevar==2, 
            CapkV{ii,1} = BuskV{ii,1}*sqrt(3);
        end
        CapKvar{ii,1} = 2*pi*handles.Freq*Cap_cuf*str2double(CapPhaseVar)*((CapkV{ii,1}*1000).^2)/1000;
    end

    [token, remain] = strtok(handles.DSSObj.ActiveCircuit.ActiveDSSElement.Properties('bus2').Val,'.');
    if strcmp(remain,'.0.0.0')|strcmp(remain,'.0.0')|strcmp(remain,'.0')
        CapGround{ii,1} = 'Solidly Grounded';
    else
        CapGround{ii,1} = 'Ungrounded or Reactor Grounded';
    end
    
    iCap = handles.DSSObj.ActiveCircuit.Capacitors.Next;
    ii=ii+1;
end

if exist('CapStatus')
    handles.CapDataTable=[CapNames,CapBusvar,CapStatus,CapKvar,FilterStatus,CapHarm,CapConn,CapPhase,CapGround,BuskV,CapkV];
else
    asdf=['','',[],'','','','','','','',''];
    handles.CapDataTable=['','',[],'','','','','','','',''];
end
set(handles.CapData,'Data',handles.CapDataTable);

Vsourceind = find(strcmp(ElementNames_mod,'Vsource'));
VsourceName = ElementNames(Vsourceind(1));
handles.DSSObj.ActiveCircuit.SetActiveElement(char(VsourceName));
Isc = handles.DSSObj.ActiveCircuit.ActiveDSSElement.Properties('Isc3').Val;
handles.Isc = Isc;
set(hand.Isc,'String',Isc)
Isc1 = handles.DSSObj.ActiveCircuit.ActiveDSSElement.Properties('Isc1').Val;
handles.Isc1 = Isc1;
set(hand.Isc1,'String',Isc1)

handles.ElementNames=ElementNames;
set(handles.CompileCircuit,'Value',0)
guidata(hObject,handles)


% --- Executes on selection change in ScanType.
function ScanType_Callback(hObject, eventdata, handles)
user_entry = get(hObject,'string');
number = get(hObject,'Value');
handles.ScanType = user_entry{number};
guidata(hObject,handles)


% --- Executes on selection change in FSphase.
function FSphase_Callback(hObject, eventdata, handles)
user_entry = get(hObject,'string');
number = get(hObject,'Value');
handles.FSphase = user_entry{number};
guidata(hObject,handles)


% --- Executes on button press in AllCapConfigs.
function AllCapConfigs_Callback(hObject, eventdata, handles)
number = get(hObject,'Value');
handles.AllCapConfigs = number;
guidata(hObject,handles)

% --- Executes on button press in AllCapConfigs2.
function AllCapConfigs2_Callback(hObject, eventdata, handles)
number = get(hObject,'Value');
handles.AllCapConfigs2 = number;
guidata(hObject,handles)


% --- Executes on button press in RunScan.
function RunScan_Callback(hObject, eventdata, handles)
k=0;
CapData=get(handles.CapData,'Data');
CapNames = CapData(find(cell2mat(CapData(:,3))),1); %This will only look at all combinations of enabled caps
% CapNames=handles.CapNames;
CapConfigStatus=zeros(length(CapNames),2^length(CapNames));

% This removes all existing harmonic sources  
handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
handles.DSSText.command = ['set loadmult=',num2str(handles.loadmult)];
handles.DSSText.command = ['batchedit load..* %SeriesRL=',num2str(handles.SeriesRL)];
handles.DSSText.command = ['batchedit vsource..* isc3=',num2str(handles.Isc)];
handles.DSSText.command = ['batchedit vsource..* isc1=',num2str(handles.Isc1)];
AddedCapData=get(handles.AddedCapacitors,'Data');
if cell2mat(AddedCapData(1,1))~=0
    for ii=1:size(AddedCapData,1)
        if cell2mat(AddedCapData(ii,3))
            handles.DSSText.Command = ['New Capacitor.',char(AddedCapData(ii,1)),' bus1=',char(AddedCapData(ii,1)),' kvar=1'];
        end
    end
end
handles.DSSText.command = 'Solve mode=snap';

% if get(handles.UseNLload,'value')==0
    handles.DSSText.command = 'batchedit generator..* Spectrum=defaultvsource';
    handles.DSSText.command = ['batchedit pvsystem..* Spectrum=defaultvsource'];
    handles.DSSText.command = ['batchedit load..* Spectrum=defaultvsource'];
    handles.DSSText.command = ['batchedit vsource..* Spectrum=defaultvsource'];
    handles.DSSText.command = ['batchedit isource..* Spectrum=defaultvsource'];
    handles.DSSText.command = ['batchedit storage..* Spectrum=defaultvsource'];
% end
MonNumber=length(handles.MonitoredElement{1});
ElementNames = handles.ElementNames;
countdata = 0;
for jj=1:MonNumber
    for ii=1:length(ElementNames)
        if strcmp(char(ElementNames(ii)),char(handles.MonitoredElement{1}{jj}))
            countdata=countdata+1;
        end
    end
end
if countdata<MonNumber
    msgbox({'The "Response Element" may not exist.'},'Warning','modal')
    return
end
Allbusnames = handles.DSSObj.ActiveCircuit.AllBusNames;
countdata=0;
for ii=1:length(Allbusnames)
    if strcmp(char(Allbusnames(ii)),char(handles.ScanBus))
        countdata=1;
    end
end
if countdata<1
    msgbox({'The "Scan Bus" may not exist.'},'Warning','modal')
    return
end

for Monii=1:MonNumber
    handles.DSSText.Command = ['New monitor.VI_',num2str(Monii),' element=',char(handles.MonitoredElement{1}{Monii}),' mode=0'];
end


if isempty(CapData)==0
    for ii=1:size(CapData,1)
        if isempty(strfind(char(CapData(ii,8)),'.1.2.3'))==0
            CapPhase=3;
        elseif isempty(strfind(char(CapData(ii,8)),'.1.2'))==0||isempty(strfind(char(CapData(ii,8)),'.2.3'))==0||isempty(strfind(char(CapData(ii,8)),'.1.3'))==0
            CapPhase=2;
        else
            CapPhase=1;
        end

        if cell2mat(CapData(ii,5))
            handles.DSSText.Command = ['Edit ',char(CapData(ii,1)),' kV=',num2str(cell2mat(CapData(ii,11))),' kvar=',num2str(cell2mat(CapData(ii,4))),' states=',num2str(cell2mat(CapData(ii,3))),' harm=',num2str(cell2mat(CapData(ii,6))),' bus1=',[strtok(char(CapData(ii,2)),'.'),char(CapData(ii,8))],' phases=',num2str(CapPhase),' conn=',char(CapData(ii,7))];
        else
            handles.DSSText.Command = ['Edit ',char(CapData(ii,1)),' kV=',num2str(cell2mat(CapData(ii,11))),' kvar=',num2str(cell2mat(CapData(ii,4))),' states=',num2str(cell2mat(CapData(ii,3))),' bus1=',[strtok(char(CapData(ii,2)),'.'),char(CapData(ii,8))],' phases=',num2str(CapPhase),' conn=',char(CapData(ii,7))];
        end
    end
end
handles.DSSText.command = ['set DataPath = "',handles.mydir,'\Export"'];
if handles.AllCapConfigs==0
    handles.DSSText.command = 'Solve mode=snap controlmode=off';
    handles.DSSText.Command = ['New spectrum.Scanspec numharm=288 harmonic=(1.083333333,1.166666667,1.25,1.333333333,1.416666667,1.5,1.583333333,1.666666667,1.75,1.833333333,1.916666667,2,2.083333333,2.166666667,2.25,2.333333333,2.416666667,2.5,2.583333333,2.666666667,2.75,2.833333333,2.916666667,3,3.083333333,3.166666667,3.25,3.333333333,3.416666667,3.5,3.583333333,3.666666667,3.75,3.833333333,3.916666667,4,4.083333333,4.166666667,4.25,4.333333333,4.416666667,4.5,4.583333333,4.666666667,4.75,4.833333333,4.916666667,5,5.083333333,5.166666667,5.25,5.333333333,5.416666667,5.5,5.583333333,5.666666667,5.75,5.833333333,5.916666667,6,6.083333333,6.166666667,6.25,6.333333333,6.416666667,6.5,6.583333333,6.666666667,6.75,6.833333333,6.916666667,7,7.083333333,7.166666667,7.25,7.333333333,7.416666667,7.5,7.583333333,7.666666667,7.75,7.833333333,7.916666667,8,8.083333333,8.166666667,8.25,8.333333333,8.416666667,8.5,8.583333333,8.666666667,8.75,8.833333333,8.916666667,9,9.083333333,9.166666667,9.25,9.333333333,9.416666667,9.5,9.583333333,9.666666667,9.75,9.833333333,9.916666667,10,10.08333333,10.16666667,10.25,10.33333333,10.41666667,10.5,10.58333333,10.66666667,10.75,10.83333333,10.91666667,11,11.08333333,11.16666667,11.25,11.33333333,11.41666667,11.5,11.58333333,11.66666667,11.75,11.83333333,11.91666667,12,12.08333333,12.16666667,12.25,12.33333333,12.41666667,12.5,12.58333333,12.66666667,12.75,12.83333333,12.91666667,13,13.08333333,13.16666667,13.25,13.33333333,13.41666667,13.5,13.58333333,13.66666667,13.75,13.83333333,13.91666667,14,14.08333333,14.16666667,14.25,14.33333333,14.41666667,14.5,14.58333333,14.66666667,14.75,14.83333333,14.91666667,15,15.08333333,15.16666667,15.25,15.33333333,15.41666667,15.5,15.58333333,15.66666667,15.75,15.83333333,15.91666667,16,16.08333333,16.16666667,16.25,16.33333333,16.41666667,16.5,16.58333333,16.66666667,16.74999999,16.83333332,16.91666665,16.99999998,17.08333331,17.16666664,17.24999997,17.3333333,17.41666663,17.49999996,17.58333329,17.66666662,17.74999995,17.83333328,17.91666661,17.99999994,18.08333327,18.1666666,18.24999993,18.33333326,18.41666659,18.49999992,18.58333325,18.66666658,18.74999991,18.83333324,18.91666657,18.9999999,19.08333323,19.16666656,19.24999989,19.33333322,19.41666655,19.49999988,19.58333321,19.66666654,19.74999987,19.8333332,19.91666653,19.99999986,20.08333319,20.16666652,20.24999985,20.33333318,20.41666651,20.49999984,20.58333317,20.6666665,20.74999983,20.83333316,20.91666649,20.99999982,21.08333315,21.16666648,21.24999981,21.33333314,21.41666647,21.4999998,21.58333313,21.66666646,21.74999979,21.83333312,21.91666645,21.99999978,22.08333311,22.16666644,22.24999977,22.3333331,22.41666643,22.49999976,22.58333309,22.66666642,22.74999975,22.83333308,22.91666641,22.99999974,23.08333307,23.1666664,23.24999973,23.33333306,23.41666639,23.49999972,23.58333305,23.66666638,23.74999971,23.83333304,23.91666637,23.9999997,24.08333303,24.16666636,24.24999969,24.33333302,24.41666635,24.49999968,24.58333301,24.66666634,24.74999967,24.833333,24.91666633,24.99999966) %mag=(100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100) angle=(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)'];
    handles.DSSText.Command = ['New Isource.scansource bus1=',handles.ScanBus,' amps=1 spectrum=scanspec scantype=',handles.ScanType];
    handles.DSSText.command = ['set casename=ResonanceScan_',num2str(k)];
    handles.DSSText.command = 'Solve mode=harmonic';
    for Monii=1:MonNumber
        handles.DSSText.command = ['export monitor VI_',num2str(Monii)];
    end
elseif handles.AllCapConfigs==1
    h = waitbar(0,'Please wait...');
    for ii=0:1:length(CapNames) %This if for th nchoosek as in how many caps to choose enabled.
%         waitbar((ii / length(CapNames)),h)
        if ii==0
            handles.DSSText.command = 'Solve mode=snap controlmode=off';
            handles.DSSText.Command = ['New spectrum.Scanspec numharm=288 harmonic=(1.083333333,1.166666667,1.25,1.333333333,1.416666667,1.5,1.583333333,1.666666667,1.75,1.833333333,1.916666667,2,2.083333333,2.166666667,2.25,2.333333333,2.416666667,2.5,2.583333333,2.666666667,2.75,2.833333333,2.916666667,3,3.083333333,3.166666667,3.25,3.333333333,3.416666667,3.5,3.583333333,3.666666667,3.75,3.833333333,3.916666667,4,4.083333333,4.166666667,4.25,4.333333333,4.416666667,4.5,4.583333333,4.666666667,4.75,4.833333333,4.916666667,5,5.083333333,5.166666667,5.25,5.333333333,5.416666667,5.5,5.583333333,5.666666667,5.75,5.833333333,5.916666667,6,6.083333333,6.166666667,6.25,6.333333333,6.416666667,6.5,6.583333333,6.666666667,6.75,6.833333333,6.916666667,7,7.083333333,7.166666667,7.25,7.333333333,7.416666667,7.5,7.583333333,7.666666667,7.75,7.833333333,7.916666667,8,8.083333333,8.166666667,8.25,8.333333333,8.416666667,8.5,8.583333333,8.666666667,8.75,8.833333333,8.916666667,9,9.083333333,9.166666667,9.25,9.333333333,9.416666667,9.5,9.583333333,9.666666667,9.75,9.833333333,9.916666667,10,10.08333333,10.16666667,10.25,10.33333333,10.41666667,10.5,10.58333333,10.66666667,10.75,10.83333333,10.91666667,11,11.08333333,11.16666667,11.25,11.33333333,11.41666667,11.5,11.58333333,11.66666667,11.75,11.83333333,11.91666667,12,12.08333333,12.16666667,12.25,12.33333333,12.41666667,12.5,12.58333333,12.66666667,12.75,12.83333333,12.91666667,13,13.08333333,13.16666667,13.25,13.33333333,13.41666667,13.5,13.58333333,13.66666667,13.75,13.83333333,13.91666667,14,14.08333333,14.16666667,14.25,14.33333333,14.41666667,14.5,14.58333333,14.66666667,14.75,14.83333333,14.91666667,15,15.08333333,15.16666667,15.25,15.33333333,15.41666667,15.5,15.58333333,15.66666667,15.75,15.83333333,15.91666667,16,16.08333333,16.16666667,16.25,16.33333333,16.41666667,16.5,16.58333333,16.66666667,16.74999999,16.83333332,16.91666665,16.99999998,17.08333331,17.16666664,17.24999997,17.3333333,17.41666663,17.49999996,17.58333329,17.66666662,17.74999995,17.83333328,17.91666661,17.99999994,18.08333327,18.1666666,18.24999993,18.33333326,18.41666659,18.49999992,18.58333325,18.66666658,18.74999991,18.83333324,18.91666657,18.9999999,19.08333323,19.16666656,19.24999989,19.33333322,19.41666655,19.49999988,19.58333321,19.66666654,19.74999987,19.8333332,19.91666653,19.99999986,20.08333319,20.16666652,20.24999985,20.33333318,20.41666651,20.49999984,20.58333317,20.6666665,20.74999983,20.83333316,20.91666649,20.99999982,21.08333315,21.16666648,21.24999981,21.33333314,21.41666647,21.4999998,21.58333313,21.66666646,21.74999979,21.83333312,21.91666645,21.99999978,22.08333311,22.16666644,22.24999977,22.3333331,22.41666643,22.49999976,22.58333309,22.66666642,22.74999975,22.83333308,22.91666641,22.99999974,23.08333307,23.1666664,23.24999973,23.33333306,23.41666639,23.49999972,23.58333305,23.66666638,23.74999971,23.83333304,23.91666637,23.9999997,24.08333303,24.16666636,24.24999969,24.33333302,24.41666635,24.49999968,24.58333301,24.66666634,24.74999967,24.833333,24.91666633,24.99999966) %mag=(100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100) angle=(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)'];
            handles.DSSText.Command = ['New Isource.scansource bus1=',handles.ScanBus,' amps=1 spectrum=scanspec scantype=',handles.ScanType];
            handles.DSSText.command = ['set casename=ResonanceScan_',num2str(k)];
            for kk=1:1:length(CapNames)
                handles.DSSText.command = ['edit ',CapNames{kk},' states=0'];
            end
            handles.DSSText.command = 'Solve mode=harmonic';
            for Monii=1:MonNumber
                handles.DSSText.command = ['export monitor VI_',num2str(Monii)];
                var1 = importdata([handles.mydir,'\Export\ResonanceScan_',num2str(k),'_Mon_vi_',num2str(Monii),'.csv'],',',1);
                if length(var1.data)<289
                    close(h)
                    return
                end
            end
            k=k+1;
            waitbar((k / (2^length(CapNames))),h,['Please wait... Solution ',num2str(k),' of ',num2str(2^length(CapNames))])
        else
            C=nchoosek(1:1:length(CapNames),ii);
            for jj=1:size(C,1)
                C(jj,:);
                handles.DSSText.command = 'Solve mode=snap controlmode=off';
                handles.DSSText.command = ['set casename=ResonanceScan_',num2str(k)];
                for kk=1:1:length(CapNames)
                    handles.DSSText.command = ['edit ',CapNames{kk},' states=0'];
                end
                for kk=1:1:size(C,2)
                    handles.DSSText.command = ['edit ',CapNames{C(jj,kk)},' states=1'];
                    CapConfigStatus(C(jj,kk),k+1)=1;
                end
                handles.DSSText.command = 'Solve mode=harmonic';
                for Monii=1:MonNumber
                    handles.DSSText.command = ['export monitor VI_',num2str(Monii)];
                    var1 = importdata([handles.mydir,'\Export\ResonanceScan_',num2str(k),'_Mon_vi_',num2str(Monii),'.csv'],',',1);
                    if length(var1.data)<289
                        close(h)
                        return
                    end
                end
                k=k+1;
                waitbar((k / (2^length(CapNames))),h,['Please wait... Solution ',num2str(k),' of ',num2str(2^length(CapNames))])
            end
        end
    end
    close(h)
end
axes(handles.FSaxes)
plot(0,0)
handles.CapNamesAnalyzed=CapNames;
handles.CapConfigStatus=CapConfigStatus;
guidata(hObject,handles)


% --- Executes on button press in LoadScanResults.
function LoadScanResults_Callback(hObject, eventdata, handles)
axes(handles.FSaxes)
plot(0,0)
MonNumber=length(handles.MonitoredElement{1});
CapNames = handles.CapNamesAnalyzed; %This will only look at all combinations of enabled caps analyzed in the RunScan
% CapNames=handles.CapNames; % Replace the variable in this function

figure(100)
close(100)
inter = figure(100);
% plotbrowser(inter,'on')

for Monii=1:MonNumber
    if Monii==1,plotcolor='b';end
    if Monii==2,plotcolor='g';end
    if Monii==3,plotcolor='r';end
    if Monii==4,plotcolor='c';end
    if Monii==5,plotcolor='m';end
    if Monii==6,plotcolor='k';end
    
    
    WCdata=zeros(12,3,2^length(CapNames));
    
    var1 = importdata([handles.mydir,'\Export\ResonanceScan_0_Mon_vi_',num2str(Monii),'.csv'],',',1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V1'}));
    [null,indHarm]=find(strcmp(var1.colheaders,{' Harmonic'}));
    if strcmp(handles.FSphase,'1')
%         figure(4)
        axes(handles.FSaxes), hold on
        aa(Monii)=plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor);
        figure(inter)
        hold on
        caps = strcat('1-',handles.MonitoredElement{1,1}{Monii},'-NoCaps');
        plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor,'displayname',caps)
    end
    WCdata(:,1,1)=[var1.data(25,indV1),var1.data(49,indV1),var1.data(73,indV1),var1.data(97,indV1),var1.data(121,indV1),var1.data(145,indV1),var1.data(169,indV1),var1.data(193,indV1),var1.data(217,indV1),var1.data(241,indV1),var1.data(265,indV1),var1.data(289,indV1)]./var1.data(1,indV1)*100;
    if handles.AllCapConfigs==1
        TotalCaps=length(CapNames);
        for kk=1:2^TotalCaps-1
            var1 = importdata([handles.mydir,'\Export\ResonanceScan_',num2str(kk),'_Mon_vi_',num2str(Monii),'.csv'],',',1);
            if strcmp(handles.FSphase,'1')
                axes(handles.FSaxes)
                hold on
                plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor)
                figure(inter)
                hold on
                ind=find(handles.CapConfigStatus(:,kk+1));
                caps=strcat(num2str(kk+1),'-',handles.MonitoredElement{1,1}{Monii},'-');
                for ii=1:length(ind)
                    caps = char(strcat(caps,CapNames(ind(ii)),','));
                end
                plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor,'displayname',caps)
            end
            WCdata(:,1,kk+1)=[var1.data(25,indV1),var1.data(49,indV1),var1.data(73,indV1),var1.data(97,indV1),var1.data(121,indV1),var1.data(145,indV1),var1.data(169,indV1),var1.data(193,indV1),var1.data(217,indV1),var1.data(241,indV1),var1.data(265,indV1),var1.data(289,indV1)]./var1.data(1,indV1)*100;
        end
    end
    if strcmp(handles.FSphase,'1')
        axes(handles.FSaxes)
        xlabel('Harmonic')
        ylabel('Voltage (V)')
        title(['Phase 1 ',handles.ScanType,' Sequence - Bus ',handles.ScanBus])
        set(gca,'XTick',[1:2:25])
        grid on
        if Monii==MonNumber,legend(aa,handles.MonitoredElement{1});end
        hold off;
        figure(inter)
        xlabel('Harmonic')
        ylabel('Voltage (V)')
%         legend('show','location','EastOutside')
        title(['Phase 1 ',handles.ScanType,' Sequence - Bus ',handles.ScanBus])
        set(gca,'XTick',[1:2:25])
        grid on
    end
    
    var1 = importdata([handles.mydir,'\Export\ResonanceScan_0_Mon_vi_',num2str(Monii),'.csv'],',',1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V2'}));
    if isempty(indV1)==0
        if strcmp(handles.FSphase,'2')
            axes(handles.FSaxes), hold on
            ab(Monii)=plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor);
            figure(inter)
            hold on
            caps = strcat('1-',handles.MonitoredElement{1,1}{Monii},' -No Caps');
            plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor,'displayname',caps)
        end
        WCdata(:,2,1)=[var1.data(25,indV1),var1.data(49,indV1),var1.data(73,indV1),var1.data(97,indV1),var1.data(121,indV1),var1.data(145,indV1),var1.data(169,indV1),var1.data(193,indV1),var1.data(217,indV1),var1.data(241,indV1),var1.data(265,indV1),var1.data(289,indV1)]./var1.data(1,indV1)*100;
        if handles.AllCapConfigs==1
            TotalCaps=length(CapNames);
            for kk=1:2^TotalCaps-1
                var1 = importdata([handles.mydir,'\Export\ResonanceScan_',num2str(kk),'_Mon_vi_',num2str(Monii),'.csv'],',',1);
                if strcmp(handles.FSphase,'2')
                    axes(handles.FSaxes)
                    hold on
                    plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor)
                    figure(inter)
                    hold on
                    ind=find(handles.CapConfigStatus(:,kk+1));
                    caps=strcat(num2str(kk+1),'-',handles.MonitoredElement{1,1}{Monii},'-');
                    for ii=1:length(ind)
                        caps = char(strcat(caps,CapNames(ind(ii)),','));
                    end
                    plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor,'displayname',caps)
                end
                WCdata(:,2,kk+1)=[var1.data(25,indV1),var1.data(49,indV1),var1.data(73,indV1),var1.data(97,indV1),var1.data(121,indV1),var1.data(145,indV1),var1.data(169,indV1),var1.data(193,indV1),var1.data(217,indV1),var1.data(241,indV1),var1.data(265,indV1),var1.data(289,indV1)]./var1.data(1,indV1)*100;
            end
        end
        if strcmp(handles.FSphase,'2')
            axes(handles.FSaxes)
            xlabel('Harmonic')
            ylabel('Voltage (V)')
            title(['Phase 2 ',handles.ScanType,' Sequence - Bus ',handles.ScanBus])
            set(gca,'XTick',[1:2:25])
            grid on
            if Monii==MonNumber,legend(ab,handles.MonitoredElement{1});end
            hold off;
            figure(inter)
            xlabel('Harmonic')
            ylabel('Voltage (V)')
            title(['Phase 2 ',handles.ScanType,' Sequence - Bus ',handles.ScanBus])
            set(gca,'XTick',[1:2:25])
            grid on
        end
    end
    
    var1 = importdata([handles.mydir,'\Export\ResonanceScan_0_Mon_vi_',num2str(Monii),'.csv'],',',1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V3'}));
    if isempty(indV1)==0
        if strcmp(handles.FSphase,'3')
            axes(handles.FSaxes), hold on
            ac(Monii)=plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor);
            figure(inter)
            hold on
            caps = strcat('1-',handles.MonitoredElement{1,1}{Monii},' -No Caps');
            plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor,'displayname',caps)
        end
        WCdata(:,3,1)=[var1.data(25,indV1),var1.data(49,indV1),var1.data(73,indV1),var1.data(97,indV1),var1.data(121,indV1),var1.data(145,indV1),var1.data(169,indV1),var1.data(193,indV1),var1.data(217,indV1),var1.data(241,indV1),var1.data(265,indV1),var1.data(289,indV1)]./var1.data(1,indV1)*100;
        if handles.AllCapConfigs==1
            TotalCaps=length(CapNames);
            for kk=1:2^TotalCaps-1
                var1 = importdata([handles.mydir,'\Export\ResonanceScan_',num2str(kk),'_Mon_vi_',num2str(Monii),'.csv'],',',1);
                if strcmp(handles.FSphase,'3')
                    axes(handles.FSaxes)
                    hold on
                    plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor)
                    figure(inter)
                    hold on
                    ind=find(handles.CapConfigStatus(:,kk+1));
                    caps=strcat(num2str(kk+1),'-',handles.MonitoredElement{1,1}{Monii},'-');
                    for ii=1:length(ind)
                        caps = char(strcat(caps,CapNames(ind(ii)),','));
                    end
                    plot(var1.data(2:size(var1.data,1)-1,indHarm),var1.data(2:size(var1.data,1)-1,indV1),plotcolor,'displayname',caps)
                end
                WCdata(:,3,kk+1)=[var1.data(25,indV1),var1.data(49,indV1),var1.data(73,indV1),var1.data(97,indV1),var1.data(121,indV1),var1.data(145,indV1),var1.data(169,indV1),var1.data(193,indV1),var1.data(217,indV1),var1.data(241,indV1),var1.data(265,indV1),var1.data(289,indV1)]./var1.data(1,indV1)*100;
            end
        end
        if strcmp(handles.FSphase,'3')
            axes(handles.FSaxes)
            xlabel('Harmonic')
            ylabel('Voltage (V)')
            title(['Phase 3 ',handles.ScanType,' Sequence - Bus ',handles.ScanBus])
            set(gca,'XTick',[1:2:25])
            grid on
            if Monii==MonNumber,legend(ac,handles.MonitoredElement{1});end
            hold off;

            figure(inter)
            xlabel('Harmonic')
            ylabel('Voltage (V)')
            title(['Phase 3 ',handles.ScanType,' Sequence - Bus ',handles.ScanBus])
            set(gca,'XTick',[1:2:25])
            grid on
        end
    end
    if handles.AllCapConfigs==1
        WCdata_var=max(WCdata,[],2);
        Config=[find(WCdata_var(1,1,:)==max(squeeze(WCdata_var(1,1,:)))),...
            find(WCdata_var(2,1,:)==max(squeeze(WCdata_var(2,1,:)))),...
            find(WCdata_var(3,1,:)==max(squeeze(WCdata_var(3,1,:)))),...
            find(WCdata_var(4,1,:)==max(squeeze(WCdata_var(4,1,:)))),...
            find(WCdata_var(5,1,:)==max(squeeze(WCdata_var(5,1,:)))),...
            find(WCdata_var(6,1,:)==max(squeeze(WCdata_var(6,1,:)))),...
            find(WCdata_var(7,1,:)==max(squeeze(WCdata_var(7,1,:)))),...
            find(WCdata_var(8,1,:)==max(squeeze(WCdata_var(8,1,:)))),...
            find(WCdata_var(9,1,:)==max(squeeze(WCdata_var(9,1,:)))),...
            find(WCdata_var(10,1,:)==max(squeeze(WCdata_var(10,1,:)))),...
            find(WCdata_var(11,1,:)==max(squeeze(WCdata_var(11,1,:)))),...
            find(WCdata_var(12,1,:)==max(squeeze(WCdata_var(12,1,:))))];
        %     WorseCaseConfig=zeros(length(handles.CapNames),12);
        CapConfigStatus=handles.CapConfigStatus;
        WorseCaseConfig=[CapConfigStatus(:,Config(1)),...
            CapConfigStatus(:,Config(2)),...
            CapConfigStatus(:,Config(3)),...
            CapConfigStatus(:,Config(4)),...
            CapConfigStatus(:,Config(5)),...
            CapConfigStatus(:,Config(6)),...
            CapConfigStatus(:,Config(7)),...
            CapConfigStatus(:,Config(8)),...
            CapConfigStatus(:,Config(9)),...
            CapConfigStatus(:,Config(10)),...
            CapConfigStatus(:,Config(11)),...
            CapConfigStatus(:,Config(12))];
        f = figure(20+Monii);
        set(f,'Name',['Cap Status for worst harmonic magnification at ',handles.MonitoredElement{1}{Monii}],'Position',[200 200 1000 150],'Resize','off');
        cnames = {'H3','H5','H7','H9','H11','H13','H15','H17','H19','H21','H23','H25'};
        rnames = CapNames;
        uitable('Parent',f,'Data',WorseCaseConfig,'ColumnName',cnames,...
            'RowName',rnames,'Position',[20 20 960 110]);
    end
end
print(['-f100'],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/07_Resonance_ScanType-',handles.ScanType,'_Phase-',handles.FSphase,'_ScanBus-',handles.ScanBus,'_AllCaps-',num2str(handles.AllCapConfigs)])



% --- Executes on selection change in LoadSpectrum.
function LoadSpectrum_Callback(hObject, eventdata, handles)
user_entry = get(hObject,'string');
number = get(hObject,'Value');
handles.LoadSpectrum = user_entry{number};
guidata(hObject,handles)
eventdata = [];
PQ_gui('DistortionTable_CellEditCallback',hObject,eventdata,guidata(hObject))

% --- Executes on button press in SaveLoadSpectrum.
function SaveLoadSpectrum_Callback(hObject, eventdata, handles)
DistortionTable=get(handles.DistortionTable,'Data');
% uisave('DistortionTable','UserDefinedSpectrum.mat')
[file,path] = uiputfile('*.txt','Save file name');
save([path,file],'DistortionTable','-ASCII')


% --- Executes when entered data in editable cell(s) in DistortionTable.
function DistortionTable_CellEditCallback(hObject, eventdata, handles)
if isempty(eventdata)
    if strcmp(handles.LoadSpectrum,'IdealLoad')
        DistortionTable=[100,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(handles.LoadSpectrum,'BackGroundHeavyWithAngles')
        DistortionTable=[100,0;0,0;8.4,148.6;0,0;4.67,-78.34;0,0;2.85,199.73;0,0;2.87,83.23;0,0;1.06,-76.3;0,0;0.91,158.74;0,0;0.94,-35.72;0,0;0.63,157.68;0,0;0.45,18.45;0,0;0.23,-81.78;0,0;0.23,89.55;0,0;0.17,-57.27];
    elseif strcmp(handles.LoadSpectrum,'BackGroundHeavyWithoutAngles')
        DistortionTable=[100,0;0,0;8.4,0;0,0;4.67,0;0,0;2.85,0;0,0;2.87,0;0,0;1.06,0;0,0;0.91,0;0,0;0.94,0;0,0;0.63,0;0,0;0.45,0;0,0;0.23,0;0,0;0.23,0;0,0;0.17,0];
    elseif strcmp(handles.LoadSpectrum,'OpenSaved')
%         uiopen('load')
        [file,path] = uigetfile('*.txt','Save file name');
        DistortionTable = load([path,file],'-ASCII');
    end
    BlankTable = zeros(100,2);
    BlankTable(1:size(DistortionTable,1),:) = DistortionTable;
    set(handles.DistortionTable,'Data',BlankTable)
end
guidata(hObject,handles)

% --- Executes on selection change in SourceSpectrum.
function SourceSpectrum_Callback(hObject, eventdata, handles)
user_entry = get(hObject,'string');
number = get(hObject,'Value');
handles.SourceSpectrum = user_entry{number};
guidata(hObject,handles)
eventdata = [];
PQ_gui('SourceDistortionTable_CellEditCallback',hObject,eventdata,guidata(hObject))

% --- Executes on button press in SaveSourceSpectrum.
function SaveSourceSpectrum_Callback(hObject, eventdata, handles)
DistortionTable=get(handles.SourceDistortionTable,'Data');
% uisave('DistortionTable','UserDefinedSpectrum.mat')
[file,path] = uiputfile('*.txt','Save file name');
save([path,file],'DistortionTable','-ASCII')

% --- Executes when entered data in editable cell(s) in Sourcecurrent distortionTable.
function SourceDistortionTable_CellEditCallback(hObject, eventdata, handles)
if isempty(eventdata)
    if strcmp(handles.SourceSpectrum,'Default')
        SourceDistortionTable=[100,0;0,0;33,0;0,0;20,0;0,0;14,0;0,0;11,0;0,0;9,0;0,0;7,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(handles.SourceSpectrum,'DefaultGen')
        SourceDistortionTable=[100,0;0,0;5,0;0,0;3,0;0,0;1.5,0;0,0;1,0;0,0;0.7,0;0,0;0.5,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(handles.SourceSpectrum,'DefaultLoad')
        SourceDistortionTable=[100,0;0,0;1.5,180;0,0;20,180;0,0;14,180;0,0;1,180;0,0;9,180;0,0;7,180;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(handles.SourceSpectrum,'IdealVsource')
        SourceDistortionTable=[100,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(handles.SourceSpectrum,'HarmonicVsource')
        SourceDistortionTable=[100,0;0,0;1.3,0;0,0;1.5,0;0,0;0.4,0;0,0;0.2,0;0,0;0.1,0;0,0;0.1,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(handles.SourceSpectrum,'OpenSaved')
%         uiopen('load') 
        [file,path] = uigetfile('*.txt','Open file name');
        SourceDistortionTable = load([path,file],'-ASCII');
    end
    BlankTable = zeros(100,2);
    BlankTable(1:size(SourceDistortionTable,1),:) = SourceDistortionTable;
    set(handles.SourceDistortionTable,'Data',BlankTable)
end
guidata(hObject,handles)


% --- Executes when entered data in editable cell(s) in AddedElements.
function AddedElements_CellEditCallback(hObject, eventdata, handles)
handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
handles.DSSText.command = 'Solve mode=snap';
tabledata = get(handles.AddedElements,'Data');
if eventdata.Indices(2)==1
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==2
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==3
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==18
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==4
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==5
    PQ_gui('DGDistortionTable_CellEditCallback',hObject,{eventdata.NewData,eventdata.Indices(1)},guidata(hObject))
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==6
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==10
    if eventdata.NewData==1
        fid = fopen([handles.mydir,'\DSSViewIntercom.txt'],'r');
        user_entry = fgetl(fid);
        fclose(fid);
        NewLoadBus = char(handles.DSSObj.ActiveCircuit.CktElements(char(user_entry)).bus(1));
        row = eventdata.Indices(1);
        tabledata(row,4)=cellstr(NewLoadBus);
        tabledata(row,10)=[{false}];
        tabledata(row,11)=[{false}];
        set(handles.AddedElements,'Data',tabledata)
    end
end

if eventdata.Indices(2)==7
    if eventdata.NewData==1
        row = eventdata.Indices(1);
        ElementNames_mod = strtok(handles.ElementNames,'.');
        loadind = find(strcmp(ElementNames_mod,'Load'));
        iLoad = handles.DSSObj.ActiveCircuit.Loads.First;
        ii=1;
        while iLoad > 0
            LoadkW{ii,1} = handles.DSSObj.ActiveCircuit.Loads.kW;
            LoadPF{ii,1} = handles.DSSObj.ActiveCircuit.Loads.PF;
            iLoad = handles.DSSObj.ActiveCircuit.Loads.Next;
            ii=ii+1;
        end
        tt=1;
        for ii=1:length(loadind)
            if LoadkW{ii,1}>300
                % Large loads such as parallel feeder load lumped at the sub should not be
                % considered in the # of potential new load locations. One
                % customer count is used in case these large loads may
                % exist on the feeder.
                loadind_mod(tt)=loadind(ii);
                LoadkWnew(tt) = LoadkW{ii,1};
                LoadPFnew(tt) = LoadPF{ii,1};
                tt=tt+1;
            else
                for jj=1:ceil(LoadkW{ii,1}/5)
                    loadind_mod(tt)=loadind(ii);
                    LoadkWnew(tt) = LoadkW{ii,1};
                    LoadPFnew(tt) = LoadPF{ii,1};
                    tt=tt+1;
                end
            end
        end
        NewOrder = randperm(length(loadind_mod));
        LoadNames = handles.ElementNames(loadind_mod(NewOrder));
        LoadNameskW = LoadkWnew(NewOrder);
        LoadNamesPF = LoadPFnew(NewOrder);
        fid1 = fopen ([handles.mydir,'\AddedElements\',char(tabledata(row,1)),'.txt'],'w');
        fid2 = fopen ([handles.mydir,'\AddedElements\',char(tabledata(row,1)),'_LoadEdit.txt'],'w');
        ID = round(rand(1)*1000000000000);
        HarmonicNumber = size(handles.AddedLoadDistortionData,2);
        harmonics = 1:1:HarmonicNumber;
        fprintf(fid1,'%s\r\n',['New Spectrum.ID',num2str(ID),' numharm=',num2str(HarmonicNumber),' harmonic=(',num2str(harmonics),') %mag=(',num2str(handles.AddedLoadDistortionData(row,:,1)),') angle=(',num2str(handles.AddedLoadDistortionData(row,:,2)),')']);
        if isempty(cell2mat(tabledata(row,18)))
            newpf = num2str(0.95);
        elseif isnan(cell2mat(tabledata(row,18)))
            newpf = num2str(0.95);
        else
            newpf = num2str(cell2mat(tabledata(row,18)));
        end
        if strcmp(tabledata(row,6),'Load')
            if isempty(cell2mat(tabledata(row,4)))==0
                phases = handles.DSSObj.ActiveCircuit.Buses(strtok(cell2mat(tabledata(row,4)),'.')).NumNodes;
                kV = handles.DSSObj.ActiveCircuit.Buses(strtok(cell2mat(tabledata(row,4)),'.')).kVBase;
                if cell2mat(tabledata(row,11))==1
                    fprintf(fid1,'%s\r\n',['New Transformer.',num2str(cell2mat(tabledata(row,1))),' phases=',num2str(phases),' buses=(',...
                        cell2mat(tabledata(row,4)),',',cell2mat(tabledata(row,12)),') kVs=(',num2str(kV),',',num2str(cell2mat(tabledata(row,13))),') kVAs=(',...
                        num2str(cell2mat(tabledata(row,15))),',',num2str(cell2mat(tabledata(row,15))),') conns=(',cell2mat(tabledata(row,14)),') %loadloss=',...
                        num2str(cell2mat(tabledata(row,16))),' XHL=',num2str(cell2mat(tabledata(row,17)))]);
                    fprintf(fid1,'%s\r\n',['New Load.NewLoad_ID',num2str(ID),'_0 phases=',num2str(phases),' bus1=',...
                        cell2mat(tabledata(row,12)),' kV=',num2str(cell2mat(tabledata(row,13))),' kW=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' spectrum=ID',num2str(ID)]);
                else
                    fprintf(fid1,'%s\r\n',['New Load.NewLoad_ID',num2str(ID),'_0 phases=',num2str(phases),' bus1=',...
                        cell2mat(tabledata(row,4)),' kV=',num2str(kV),' kW=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' spectrum=ID',num2str(ID)]);
                end
            end            
            for ii=1:round(length(LoadNames)*cell2mat(tabledata(row,2))/100)
                bus=handles.DSSObj.ActiveCircuit.CktElements(char(LoadNames(ii))).Bus;
                phases=handles.DSSObj.ActiveCircuit.CktElements(char(LoadNames(ii))).NumPhases;
                kV=handles.DSSObj.ActiveCircuit.Buses(char(strtok(bus,'.'))).kVbase;
                fprintf(fid1,'%s\r\n',['New Load.NewLoad_ID',num2str(ID),'_',num2str(ii),' phases=',num2str(phases),' bus1=',...
                    char(bus),' kV=',num2str(kV),' kW=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' spectrum=ID',num2str(ID)]);
                if isempty(cell2mat(strfind(tabledata(row,5),'CFL')))==0 || isempty(cell2mat(strfind(tabledata(row,5),'LED')))==0
                    fprintf(fid2,'%s\r\n',[char(LoadNames(ii)),',',num2str(LoadNameskW(ii)),',',num2str(4*cell2mat(tabledata(row,3)))]);
                end
                if isempty(cell2mat(strfind(tabledata(row,5),'HVAC')))==0 
                    fprintf(fid2,'%s\r\n',[char(LoadNames(ii)),',',num2str(LoadNameskW(ii)),',',num2str(cell2mat(tabledata(row,3)))]);
                end
            end
        end
        if strcmp(tabledata(row,6),'PVsystem')
            if isempty(cell2mat(tabledata(row,4)))==0
                phases = handles.DSSObj.ActiveCircuit.Buses(strtok(cell2mat(tabledata(row,4)),'.')).NumNodes;
                kV = handles.DSSObj.ActiveCircuit.Buses(strtok(cell2mat(tabledata(row,4)),'.')).kVBase;
                pf = 0.95;
                if cell2mat(tabledata(row,11))==1
                    fprintf(fid1,'%s\r\n',['New Transformer.',num2str(cell2mat(tabledata(row,1))),' phases=',num2str(phases),' buses=(',...
                        cell2mat(tabledata(row,4)),',',cell2mat(tabledata(row,12)),') kVs=(',num2str(kV),',',num2str(cell2mat(tabledata(row,13))),') kVAs=(',...
                        num2str(cell2mat(tabledata(row,15))),',',num2str(cell2mat(tabledata(row,15))),') conns=(',cell2mat(tabledata(row,14)),') %loadloss=',...
                        num2str(cell2mat(tabledata(row,16))),' XHL=',num2str(cell2mat(tabledata(row,17)))]);
                    fprintf(fid1,'%s\r\n',['New PVsystem.NewPV_ID',num2str(ID),'_0 phases=',num2str(phases),' bus1=',...
                        cell2mat(tabledata(row,12)),' kV=',num2str(cell2mat(tabledata(row,13))),' kVA=',num2str(cell2mat(tabledata(row,3)*1.1)),' irradiance=1 Pmpp=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' model=1 Vminpu=0.5 xdp=0.5 spectrum=ID',num2str(ID)]);
                else
                    fprintf(fid1,'%s\r\n',['New PVsystem.NewPV_ID',num2str(ID),'_0 phases=',num2str(phases),' bus1=',...
                        cell2mat(tabledata(row,4)),' kV=',num2str(kV),' kVA=',num2str(cell2mat(tabledata(row,3)*1.1)),' irradiance=1 Pmpp=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' model=1 Vminpu=0.5 xdp=0.5 spectrum=ID',num2str(ID)]);
                end
            end
            for ii=1:round(length(LoadNames)*cell2mat(tabledata(row,2))/100)
                bus=handles.DSSObj.ActiveCircuit.CktElements(char(LoadNames(ii))).Bus;
                phases=handles.DSSObj.ActiveCircuit.CktElements(char(LoadNames(ii))).NumPhases;
                kV=handles.DSSObj.ActiveCircuit.Buses(char(strtok(bus,'.'))).kVbase;
                fprintf(fid1,'%s\r\n',['New PVsystem.NewPV_ID',num2str(ID),'_',num2str(ii),' phases=',num2str(phases),' bus1=',...
                    char(bus),' kV=',num2str(kV),' kVA=',num2str(cell2mat(tabledata(row,3)*1.1)),' irradiance=1 Pmpp=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' %cutin=0.1 %cutout=0.1 spectrum=ID',num2str(ID)]);
            end
        end
        if strcmp(tabledata(row,6),'Generator')
            if isempty(cell2mat(tabledata(row,4)))==0
                phases = handles.DSSObj.ActiveCircuit.Buses(strtok(cell2mat(tabledata(row,4)),'.')).NumNodes;
                kV = handles.DSSObj.ActiveCircuit.Buses(strtok(cell2mat(tabledata(row,4)),'.')).kVBase;
                pf = 0.95;
                if cell2mat(tabledata(row,11))==1
                    fprintf(fid1,'%s\r\n',['New Transformer.',num2str(cell2mat(tabledata(row,1))),' phases=',num2str(phases),' buses=(',...
                        cell2mat(tabledata(row,4)),',',cell2mat(tabledata(row,12)),') kVs=(',num2str(kV),',',num2str(cell2mat(tabledata(row,13))),') kVAs=(',...
                        num2str(cell2mat(tabledata(row,15))),',',num2str(cell2mat(tabledata(row,15))),') conns=(',cell2mat(tabledata(row,14)),') %loadloss=',...
                        num2str(cell2mat(tabledata(row,16))),' XHL=',num2str(cell2mat(tabledata(row,17)))]);
                    fprintf(fid1,'%s\r\n',['New Generator.NewGen_ID',num2str(ID),'_0 phases=',num2str(phases),' bus1=',...
                        cell2mat(tabledata(row,12)),' kV=',num2str(cell2mat(tabledata(row,13))),' kW=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' model=1 Vminpu=0.5 xdp=0.5 spectrum=ID',num2str(ID)]);
                else
                    fprintf(fid1,'%s\r\n',['New Generator.NewGen_ID',num2str(ID),'_0 phases=',num2str(phases),' bus1=',...
                        cell2mat(tabledata(row,4)),' kV=',num2str(kV),' kW=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' model=1 Vminpu=0.5 xdp=0.5 spectrum=ID',num2str(ID)]);
                end
            end            
            for ii=1:round(length(LoadNames)*cell2mat(tabledata(row,2))/100)
                bus=handles.DSSObj.ActiveCircuit.CktElements(char(LoadNames(ii))).Bus;
                phases=handles.DSSObj.ActiveCircuit.CktElements(char(LoadNames(ii))).NumPhases;
                kV=handles.DSSObj.ActiveCircuit.Buses(char(strtok(bus,'.'))).kVbase;
                fprintf(fid1,'%s\r\n',['New Generator.NewGen_ID',num2str(ID),'_',num2str(ii),' phases=',num2str(phases),' bus1=',...
                    char(bus),' kV=',num2str(kV),' kW=',num2str(cell2mat(tabledata(row,3))),' pf=',newpf,' model=1 Vminpu=0.5 xdp=0.5 spectrum=ID',num2str(ID)]);
            end
        end
        fclose(fid1);
        fclose(fid2);
        data=tabledata(row,:);
        save([handles.mydir,'\AddedElements\',char(tabledata(row,1)),'.mat'],'data')
    end
end
if eventdata.Indices(2)==8
    if eventdata.NewData==1
        oldFolder = cd(handles.mydir);
        cd([handles.mydir,'\AddedElements\'])
        uiopen('load')
        cd(oldFolder);
        row = eventdata.Indices(1);
        tabledata(row,:)=data;
        set(handles.AddedElements,'Data',tabledata)
    end
end
if eventdata.Indices(2)==12
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==13
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==14
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==15
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==16
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end
if eventdata.Indices(2)==17
    row = eventdata.Indices(1);
    tabledata(row,7:9)=[{false},{false},{false}];
    tabledata(row,11)=[{false}];
    set(handles.AddedElements,'Data',tabledata)
end


% --- Executes on button press in SaveDGSpectrum.
function SaveDGSpectrum_Callback(hObject, eventdata, handles)
DistortionTable=get(handles.DGDistortionTable,'Data');
% uisave('DistortionTable','UserDefinedSpectrum.mat')
[file,path] = uiputfile('*.txt','Save file name');
save([path,file],'DistortionTable','-ASCII')


% --- Executes when entered data in editable cell(s) in DGDistortionTable.
function DGDistortionTable_CellEditCallback(hObject, eventdata, handles)
if length(eventdata)==2
    row=cell2mat(eventdata(2));
    if  strcmp(char(eventdata(1)),'High_PF_CFL_With_Angles')
        DGDistortionTable=[100,0;0,0;30.22043355,201.8;0,0;17.1188146,33.4;0,0;5.825482484,265.3;0,0;1.747004482,101.1;0,0;1.735113875,284.6;0,0;1.680234153,67.2;0,0;1.443336687,287.3;0,0;0.655812677,143.3;0,0;0.619226196,306.5;0,0;0.46373365,138.2;0,0;0.207628281,335.7;0,0;0.254276045,216];
    elseif strcmp(char(eventdata(1)),'High_PF_CFL_Without_Angles')
        DGDistortionTable=[100,0;0,0;30.22043355,0;0,0;17.1188146,0;0,0;5.825482484,0;0,0;1.747004482,0;0,0;1.735113875,0;0,0;1.680234153,0;0,0;1.443336687,0;0,0;0.655812677,0;0,0;0.619226196,0;0,0;0.46373365,0;0,0;0.207628281,0;0,0;0.254276045,0];
    elseif  strcmp(char(eventdata(1)),'Low_PF_CFL_With_Angles')
        DGDistortionTable=[100,0;0,0;75.113,221.967;0,0;62.834,91.189;0,0;51.549,162.889;0,0;41.278,192.911;0,0;29.82,100.033;0,0;17.505,207.044;0,0;7.606,163.311;0,0;3.187,93.778;0,0;3.157,77.167;0,0;2.245,207.078;0,0;1.801,151.6;0,0;1.365,145.556];
    elseif strcmp(char(eventdata(1)),'Low_PF_CFL_Without_Angles')
        DGDistortionTable=[100,0;0,0;75.113,0;0,0;62.834,0;0,0;51.549,0;0,0;41.278,0;0,0;29.82,0;0,0;17.505,0;0,0;7.606,0;0,0;3.187,0;0,0;3.157,0;0,0;2.245,0;0,0;1.801,0;0,0;1.365,0];
     elseif  strcmp(char(eventdata(1)),'High_PF_LED_With_Angles')
        DGDistortionTable=[100,0;0,0;2.263,122.2;0,0;4.035,130.3;0,0;3.954,265.2;0,0;3.307,355.2;0,0;2.234,310.2;0,0;1.843,355.2;0,0;1.043,355.2;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'High_PF_LED_Without_Angles')
        DGDistortionTable=[100,0;0,0;2.263,0;0,0;4.035,0;0,0;3.954,0;0,0;3.307,0;0,0;2.234,0;0,0;1.843,0;0,0;1.043,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'Low_PF_LED_With_Angles')
        DGDistortionTable=[100,0;0,0;79.8,166.9;0,0;55.18,60.3;0,0;38.33,323.1;0,0;36,226.2;0,0;35.46,123.7;0,0;28.26,-10;0,0;17.94,233.1;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'Low_PF_LED_Without_Angles')
        DGDistortionTable=[100,0;0,0;79.8,0;0,0;55.18,0;0,0;38.33,0;0,0;36,0;0,0;35.46,0;0,0;28.26,0;0,0;17.94,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'AC_Charger_EV_With_Angles')
        DGDistortionTable=[100,0;0,0;8.92169448,-129.22;0,0;0.385109114,67.53;0,0;0.256739409,-157.64;0,0;0.449293967,165.9;0,0;0.513478819,172.36;0,0;0.449293967,159.91;0,0;0.513478819,159.93;0,0;0.513478819,158.69;0,0;0.449293967,143.62;0,0;0.385109114,147.34;0,0;0.385109114,140.45;0,0;0.320924262,133.7];
    elseif strcmp(char(eventdata(1)),'AC_Charger_EV_Without_Angles')
        DGDistortionTable=[100,0;0,0;8.92169448,0;0,0;0.385109114,0;0,0;0.256739409,0;0,0;0.449293967,0;0,0;0.513478819,0;0,0;0.449293967,0;0,0;0.513478819,0;0,0;0.513478819,0;0,0;0.449293967,0;0,0;0.385109114,0;0,0;0.385109114,0;0,0;0.320924262,0];
    elseif strcmp(char(eventdata(1)),'DC_Charger_EV_With_Angles')
        DGDistortionTable=[100,0;0,0;5.926179084,-200.88;0,0;2.125768968,-18.51;0,0;1.517429938,-296.16;0,0;0.827067669,-183.58;0,0;0.45112782,-99.9;0,0;1.066302119,-43.37;0,0;0.36910458,28.29;0,0;0.40328093,4.41;0,0;0.574162679,-88.69;0,0;0.11619959,-75.14;0,0;0.663021189,-42.75;0,0;0.649350649,-144.26];
    elseif strcmp(char(eventdata(1)),'DC_Charger_EV_Without_Angles')
        DGDistortionTable=[100,0;0,0;5.926179084,0;0,0;2.125768968,0;0,0;1.517429938,0;0,0;0.827067669,0;0,0;0.45112782,0;0,0;1.066302119,0;0,0;0.36910458,0;0,0;0.40328093,0;0,0;0.574162679,0;0,0;0.11619959,0;0,0;0.663021189,0;0,0;0.649350649,0];
    elseif strcmp(char(eventdata(1)),'PV_With_Angles')
        DGDistortionTable=[100,0;0,0;0.69,128.75;0,0;0.63,207.41;0,0;0.56,10.29;0,0;0.19,355.97;0,0;0.07,329.09;0,0;0.17,303.86;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'PV_Without_Angles')
        DGDistortionTable=[100,0;0,0;0.69,0;0,0;0.63,0;0,0;0.56,0;0,0;0.19,0;0,0;0.07,0;0,0;0.17,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'Home_Entertainment_With_Angles')
        DGDistortionTable=[100,0;0,0;31.132349995654,189.281666666667;0,0;21.4580758779936,194.81;0,0;14.8887603974166,184.618333333333;0,0;10.5956335574376,217.771666666667;0,0;7.89769055474304,188.166666666667;0,0;6.14567153800513,269.445;0,0;4.01498059400852,237.358333333333;0,0;2.58683567074438,197.583333333333;0,0;2.19833792501389,45.3166666666666;0,0;1.9492912183637,106.166666666667;0,0;1.63977793650254,89.8499999999999;0,0;1.33706714435261,128.785];
    elseif strcmp(char(eventdata(1)),'Home_Entertainment_Without_Angles')
        DGDistortionTable=[100,0;0,0;31.132349995654,0;0,0;21.4580758779936,0;0,0;14.8887603974166,0;0,0;10.5956335574376,0;0,0;7.89769055474304,0;0,0;6.14567153800513,0;0,0;4.01498059400852,0;0,0;2.58683567074438,0;0,0;2.19833792501389,0;0,0;1.9492912183637,0;0,0;1.63977793650254,0;0,0;1.33706714435261,0];
    elseif strcmp(char(eventdata(1)),'ECM_HVAC_With_Angles')
        DGDistortionTable=[100,0;0,0;4.88953587943141,132.32;0,0;4.28155506079808,112.8;0,0;5.41188559684877,276.23;0,0;5.60883712964549,7.39;0,0;2.29491351258777,186.17;0,0;2.36341839356054,12.472;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'ECM_HVAC_Without_Angles')
        DGDistortionTable=[100,0;0,0;4.88953587943141,0;0,0;4.28155506079808,0;0,0;5.41188559684877,0;0,0;5.60883712964549,0;0,0;2.29491351258777,0;0,0;2.36341839356054,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];    
    elseif strcmp(char(eventdata(1)),'Switch_Mode_Power_Supply_With_Angles')
        DGDistortionTable=[100,-37;0.2,65;65.7,-97;0.4,-72;37.7,-166;0.4,-154;12.7,113;0.3,112;4.4,-46;0.1,0;5.3,-158;0.1,142;2.5,92;0.1,65;1.9,-51;0,0;1.8,-151;0,0;1.1,84;0,0;0.6,-41;0,0;0.8,-148;0,0;0.4,64];
    elseif strcmp(char(eventdata(1)),'Switch_Mode_Power_Supply_Without_Angles')
        DGDistortionTable=[100,0;0.2,0;65.7,0;0.4,0;37.7,0;0.4,0;12.7,0;0.3,0;4.4,0;0.1,0;5.3,0;0.1,0;2.5,0;0.1,0;1.9,0;0,0;1.8,0;0,0;1.1,0;0,0;0.6,0;0,0;0.8,0;0,0;0.4,0];    
    elseif strcmp(char(eventdata(1)),'6-Pulse_DC_Drive_With_Angles')
        DGDistortionTable=[100,-43;0.3,68;0.4,-126;0.2,30;25.3,-30;0.1,61;5.5,-122;0.2,10;0.6,37;0.2,-45;8.2,-102;0.1,-37;3.9,170;0.2,-65;0.3,-40;0.2,-119;5,-179;0.1,-126;2.9,93;0.2,-136;0.3,-44;0.2,167;3.4,109;0.1,162;2.1,14];
    elseif strcmp(char(eventdata(1)),'6-Pulse_DC_Drive_Without_Angles')
        DGDistortionTable=[100,0;0.3,0;0.4,0;0.2,0;25.3,0;0.1,0;5.5,0;0.2,0;0.6,0;0.2,0;8.2,0;0.1,0;3.9,0;0.2,0;0.3,0;0.2,0;5,0;0.1,0;2.9,0;0.2,0;0.3,0;0.2,0;3.4,0;0.1,0;2.1,0];
    elseif strcmp(char(eventdata(1)),'Single_Phase_ASD_With_Angles')
        DGDistortionTable=[100,-14;3.8,-85;8.5,-114;3.5,-103;79.5,145;0.3,25;66,124;2.5,55;2.7,11;1.7,68;36,-92;1.2,132;21.8,-118;1.2,156;2.4,22;0.3,-136;10.4,-23;0.8,-92;8,-79;0.9,-117;1.4,131;0.5,-105;6.7,39;0,0;4.5,-2];
    elseif strcmp(char(eventdata(1)),'Single_Phase_ASD_Without_Angles')
        DGDistortionTable=[100,0;3.8,0;8.5,0;3.5,0;79.5,0;0.3,0;66,0;2.5,0;2.7,0;1.7,0;36,0;1.2,0;21.8,0;1.2,0;2.4,0;0.3,0;10.4,0;0.8,0;8,0;0.9,0;1.4,0;0.5,0;6.7,0;0,0;4.5,0];
    elseif strcmp(char(eventdata(1)),'Arc_Furnace_Initial_Melt')
        DGDistortionTable=[100,0;7.7,0;5.8,0;2.5,0;4.2,0;0,0;3.1,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0];
    elseif strcmp(char(eventdata(1)),'Six_Pulse_Converter_Ideal')
        DGDistortionTable=[0,0;0,0;0,0;0,0;20,0;0,0;14.3,0;0,0;0,0;0,0;9.1,0;0,0;7.7,0;0,0;0,0;0,0;5.9,0;0,0;0.5,0;0,0;0,0;0,0;4.3,0;0,0;4,0];
    elseif strcmp(char(eventdata(1)),'Six_Pulse_Converter_Practical')
        DGDistortionTable=[0,0;0,0;0,0;0,0;17.5,0;0,0;11,0;0,0;0,0;0,0;4.5,0;0,0;2.9,0;0,0;0,0;0,0;1.5,0;0,0;1,0;0,0;0,0;0,0;0.9,0;0,0;0.8,0];
    elseif strcmp(char(eventdata(1)),'Twelve_Pulse_Converter_Ideal')
        DGDistortionTable=[0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;9.1,0;0,0;7.7,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;4.3,0;0,0;4,0];
    elseif strcmp(char(eventdata(1)),'Twelve_Pulse_Converter_Practical')
        DGDistortionTable=[0,0;0,0;0,0;0,0;2.6,0;0,0;1.6,0;0,0;0,0;0,0;7.9,0;0,0;5.5,0;0,0;0,0;0,0;0.2,0;0,0;0.1,0;0,0;0,0;0,0;2.3,0;0,0;0.8,0];
    elseif strcmp(char(eventdata(1)),'Twentyfour_Pulse_Converter_Ideal')
        DGDistortionTable=[0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;0,0;4.3,0;0,0;4,0];
    elseif strcmp(char(eventdata(1)),'Twentyfour_Pulse_Converter_Practical')
        DGDistortionTable=[0,0;0,0;0,0;0,0;2.6,0;0,0;1.6,0;0,0;0,0;0,0;0.7,0;0,0;0.4,0;0,0;0,0;0,0;0.2,0;0,0;0.1,0;0,0;0,0;0,0;1.9,0;0,0;0.8,0];
     elseif strcmp(char(eventdata(1)),'Wind_Plant')
        DGDistortionTable=[100,0;0.343,0;0,0;0.201,0;0.435,0;0,0;0.471,0;0.401,0;0,0;0.545,0;1.45,0;0,0;1.854,0;0.475,0;0,0;0.371,0;0.761,0;0,0;0.42,0;0.324,0;0,0;0.33,0;0.37,0;0,0;0.235,0];
     elseif strcmp(char(eventdata(1)),'Load User Defined')
%         uiopen('load')
        [file,path] = uigetfile('*.txt','Open file name');
        DistortionTable = load([path,file],'-ASCII');
        DGDistortionTable=DistortionTable;
    end
    handles.AddedLoadDistortionData(row,:,:)=zeros(1,100,2);
    handles.AddedLoadDistortionData(row,1:size(DGDistortionTable,1),:)=DGDistortionTable;
    BlankTable = zeros(100,2);
    BlankTable(1:size(DGDistortionTable,1),:) = DGDistortionTable;
    set(handles.DGDistortionTable,'Data',BlankTable)
end
guidata(hObject,handles)


% --- Executes on button press in RunDistortionScan.
function RunDistortionScan_Callback(hObject, eventdata, handles)
DistortionData=get(handles.DistortionTable,'Data');
SourceDistortionTable=get(handles.SourceDistortionTable,'Data');
AddedElements=get(handles.AddedElements,'Data');

handles.DSSText.Command = ['Compile (',handles.mydir,'\Master_',handles.ckt,'.dss)'];
% handles.DSSText.Command = ['New monitor.VI element=',handles.DistortionMonitoredElement,' mode=0 residual=yes'];
% handles.DSSText.command = ['set loadmult=',num2str(handles.loadmult)];
AddedCapData=get(handles.AddedCapacitors,'Data');
if cell2mat(AddedCapData(1,1))~=0
    for ii=1:size(AddedCapData,1)
        if cell2mat(AddedCapData(ii,3))
            handles.DSSText.Command = ['New Capacitor.',char(AddedCapData(ii,1)),' bus1=',char(AddedCapData(ii,1)),' kvar=1'];
        end
    end
end
handles.DSSText.command = 'Solve mode=snap';

% set up monitors on all line segments at bus 1
ElementNames = handles.ElementNames;
ElementNames_mod = strtok(ElementNames,'.');
Lineind = find(strcmp(ElementNames_mod,'Line'));
LineNames = ElementNames(Lineind);
count=1;
for ii=1:length(LineNames)
    BusName = handles.DSSObj.ActiveCircuit.CktElements(char(LineNames(ii))).bus(1);
    BusName = strtok(BusName,'.');
    handles.DSSObj.ActiveCircuit.Buses(char(BusName));
    if handles.DSSObj.ActiveCircuit.ActiveBus.x ~= 0
        handles.DSSText.Command = ['New monitor.VI_',num2str(count),' element=',char(LineNames(ii)),' mode=0 residual=yes'];
        LineBusXcoord(count,1) = handles.DSSObj.ActiveCircuit.ActiveBus.x;
        LineBusYcoord(count,1) = handles.DSSObj.ActiveCircuit.ActiveBus.y;
        
        BusName2 = handles.DSSObj.ActiveCircuit.CktElements(char(LineNames(ii))).bus(2);
        BusName2 = strtok(BusName2,'.');
        handles.DSSObj.ActiveCircuit.Buses(char(BusName2));
        if handles.DSSObj.ActiveCircuit.ActiveBus.x ~= 0
            LineBus2Xcoord(count,1) = handles.DSSObj.ActiveCircuit.ActiveBus.x;
            LineBus2Ycoord(count,1) = handles.DSSObj.ActiveCircuit.ActiveBus.y;
            count = count+1;
        else
            LineBus2Xcoord(count,1) = LineBusXcoord(count,1);
            LineBus2Ycoord(count,1) = LineBusYcoord(count,1);
            count = count+1;
        end
    end

end
handles.LineBusXcoord=LineBusXcoord;
handles.LineBusYcoord=LineBusYcoord;
handles.LineBus2Xcoord=LineBus2Xcoord;
handles.LineBus2Ycoord=LineBus2Ycoord;
guidata(hObject,handles)

HarmonicNumber = size(DistortionData,1);
harmonics = 1:1:HarmonicNumber;
handles.DSSText.Command = ['New Spectrum.load numharm=',num2str(HarmonicNumber),' harmonic=(',num2str(harmonics),') %mag=(',num2str(transpose(DistortionData(:,1))),') angle=(',num2str(transpose(DistortionData(:,2))),')'];
HarmonicNumberSource = size(SourceDistortionTable,1);
harmonicsSource = 1:1:HarmonicNumber;
handles.DSSText.Command = ['New Spectrum.vsource numharm=',num2str(HarmonicNumberSource),' harmonic=(',num2str(harmonicsSource),') %mag=(',num2str(transpose(SourceDistortionTable(:,1))),') angle=(',num2str(transpose(SourceDistortionTable(:,2))),')'];
handles.DSSText.command = ['batchedit load..* %SeriesRL=',num2str(handles.SeriesRL)];
handles.DSSText.command = ['batchedit vsource..* isc1=',num2str(handles.Isc)];
handles.DSSText.command = ['batchedit vsource..* isc3=',num2str(handles.Isc1)];

% Remove any existing spectrums in the dss files.
if get(handles.UseNLload,'value')==0
    handles.DSSText.command = ['batchedit vsource..* Spectrum=vsource'];
    handles.DSSText.command = ['batchedit load..* Spectrum=load'];
    handles.DSSText.command = 'batchedit generator..* Spectrum=defaultvsource';
    handles.DSSText.command = ['batchedit pvsystem..* Spectrum=defaultvsource'];
    handles.DSSText.command = ['batchedit isource..* Spectrum=defaultvsource'];
    handles.DSSText.command = ['batchedit storage..* Spectrum=defaultvsource'];
end



if size(AddedElements,1)>=1
    EditNames = [];
    EditData = [];
    for ii=1:size(AddedElements,1)
        if cell2mat(AddedElements(ii,9))==1
            handles.DSSText.Command = ['Redirect (',handles.mydir,'\AddedElements\',char(AddedElements(ii,1)),'.txt)'];
            if isempty(cell2mat(strfind(AddedElements(ii,5),'CFL')))==0 || isempty(cell2mat(strfind(AddedElements(ii,5),'LED')))==0 || isempty(cell2mat(strfind(AddedElements(ii,5),'HVAC')))==0
                var1 = importdata([handles.mydir,'\AddedElements\',char(AddedElements(ii,1)),'_LoadEdit.txt'],',',0);
                if isempty(var1)==0
                    EditData=vertcat(EditData,var1.data);
                    EditNames=vertcat(EditNames,var1.textdata);
                end
            end
        end
    end
    for kk=1:length(EditNames)
        kWnew=EditData(kk,1);
        for jj=1:length(EditNames)
            if strcmp(char(EditNames(kk)),char(EditNames(jj))), kWnew = kWnew-EditData(jj,2); end
        end
        if kWnew<=0, kWnew = 0.1; end
        ID = round(rand(1)*1000000000000);
        % The next four lines adjust the existing load spectrums based on less harmonic content. If the user selects to use existing spectrums, the adjustment is ignored. 
        if get(handles.UseNLload,'value')==0
            handles.DSSText.Command = ['New Spectrum.ID',num2str(ID),' numharm=',num2str(HarmonicNumber),' harmonic=(',num2str(harmonics),') %mag=(',num2str([DistortionData(1,1),transpose(DistortionData(2:100,1).*(EditData(kk,1)/kWnew))]),') angle=(',num2str(transpose(DistortionData(:,2))),')'];
            handles.DSSText.Command = ['Edit ',char(EditNames(kk)), ' Spectrum=ID',num2str(ID)];
        end
        handles.DSSText.Command = ['Edit ',char(EditNames(kk)), ' kW=', num2str(kWnew)];
    end
end

handles.DSSText.Command = ['New monitor.VI element=',handles.DistortionMonitoredElement,' terminal=2 mode=0 residual=yes'];
handles.DSSText.Command = ['New monitor.PQ element=',handles.DistortionMonitoredElement,' terminal=2 mode=1 residual=yes ppolar=false'];
handles.DSSText.command = ['set loadmult=',num2str(handles.loadmult)];
handles.DSSText.command = 'Solve mode=snap';
% handles.DSSText.command = 'show voltages';

CapData=get(handles.CapData,'Data');
CapNames = CapData(find(cell2mat(CapData(:,3))),1); %This will only look at all combinations of enabled caps
% CapNames=handles.CapNames;
CapConfigStatus=zeros(length(CapNames),2^length(CapNames));
if isempty(CapData)==0
    for ii=1:size(CapData,1)
        if isempty(strfind(char(CapData(ii,8)),'.1.2.3'))==0
            CapPhase=3;
        elseif isempty(strfind(char(CapData(ii,8)),'.1.2'))==0||isempty(strfind(char(CapData(ii,8)),'.2.3'))==0||isempty(strfind(char(CapData(ii,8)),'.1.3'))==0
            CapPhase=2;
        else
            CapPhase=1;
        end
        
        if cell2mat(CapData(ii,5))
            handles.DSSText.Command = ['Edit ',char(CapData(ii,1)),' kV=',num2str(cell2mat(CapData(ii,11))),' kvar=',num2str(cell2mat(CapData(ii,4))),' states=',num2str(cell2mat(CapData(ii,3))),' harm=',num2str(cell2mat(CapData(ii,6))),' bus1=',[strtok(char(CapData(ii,2)),'.'),char(CapData(ii,8))],' phases=',num2str(CapPhase),' conn=',char(CapData(ii,7))];
        else
            handles.DSSText.Command = ['Edit ',char(CapData(ii,1)),' kV=',num2str(cell2mat(CapData(ii,11))),' kvar=',num2str(cell2mat(CapData(ii,4))),' states=',num2str(cell2mat(CapData(ii,3))),' bus1=',[strtok(char(CapData(ii,2)),'.'),char(CapData(ii,8))],' phases=',num2str(CapPhase),' conn=',char(CapData(ii,7))];
        end

    end
end

k=0;
handles.DSSText.command = ['set DataPath = "',handles.mydir,'\Export"'];
if handles.AllCapConfigs2==0
    handles.DSSText.command = 'Solve mode=snap controlmode=static';
    handles.DSSText.command = ['set casename=Distortion_0'];
    handles.DSSText.command = 'Solve mode=snap controlmode=off';
    handles.DSSText.command = ['Set Harmonics=(',num2str(harmonics),')'];
    handles.DSSText.command = 'Solve mode=harmonic';
    handles.DSSText.command = 'Export monitors VI';
    handles.DSSText.command = 'Export Monitors PQ';
    for ii=1:length(LineBusXcoord)
        handles.DSSText.Command = ['Export monitors VI_',num2str(ii)];
    end
elseif handles.AllCapConfigs2==1
%     CapNames=handles.CapNames;
    h = waitbar(0,'Please wait...');
    for jj=0:1:length(CapNames)
%         waitbar((jj / length(CapNames)),h)
        if jj==0
            handles.DSSText.command = ['set casename=Distortion_',num2str(k)];
            for kk=1:1:length(CapNames)
                handles.DSSText.command = ['edit ',CapNames{kk},' states=0'];
            end
            handles.DSSText.command = 'Solve mode=snap controlmode=off';
            handles.DSSText.command = 'Solve mode=harmonic';
            handles.DSSText.command = 'Export monitors VI';
            var1 = importdata([handles.mydir,'\Export\Distortion_',num2str(k),'_Mon_vi.csv'],',',1);
            if length(var1.data)<100
                close(h)
                return
            end
            for ii=1:length(LineBusXcoord)
                handles.DSSText.Command = ['Export monitors VI_',num2str(ii)];
            end
            k=k+1;
            waitbar((k / (2^length(CapNames))),h,['Please wait... Solution ',num2str(k),' of ',num2str(2^length(CapNames))])
        else
            C=nchoosek(1:1:length(CapNames),jj);
            for tt=1:size(C,1)
                C(tt,:);
                handles.DSSText.command = ['set casename=Distortion_',num2str(k)];
                for kk=1:1:length(CapNames)
                    handles.DSSText.command = ['edit ',CapNames{kk},' states=0'];
                end
                for kk=1:1:size(C,2)
                    handles.DSSText.command = ['edit ',CapNames{C(tt,kk)},' states=1'];
                    CapConfigStatus(C(tt,kk),k+1)=1;
                end
                handles.DSSText.command = 'Solve mode=snap controlmode=off';
                handles.DSSText.command = 'Solve mode=harmonic';
                handles.DSSText.command = 'Export monitors VI';
                handles.DSSText.command = 'Export Monitors PQ';
                var1 = importdata([handles.mydir,'\Export\Distortion_',num2str(k),'_Mon_vi.csv'],',',1);
                if length(var1.data)<100
                    close(h)
                    return
                end
                k=k+1;
                waitbar((k / (2^length(CapNames))),h,['Please wait... Solution ',num2str(k),' of ',num2str(2^length(CapNames))])
            end
        end
    end
    close(h)
end
handles.CapNamesAnalyzed2=CapNames;
handles.CapConfigStatus2=CapConfigStatus;

%Get Isc data
handles.DSSText.command = 'Batchedit Capacitor..* enabled=false';
handles.DSSText.command = 'Batchedit Load..* enabled=false';
handles.DSSText.command = 'Solve mode=direct';
handles.DSSText.command = ['New Fault.X bus1=',char(handles.DistortionBus)];
handles.DSSText.command = 'Solve controlmode=off mode=dynamics number=1 loadmodel=admittance';
handles.DSSText.command = ['set casename=FaultIsc'];
handles.DSSText.command = 'Export Monitors VI';
var1 = importdata([handles.mydir,'\Export\FaultIsc_Mon_vi.csv'],',',1);
handles.iscDistortionBus = var1.data(11);

guidata(hObject,handles)


function Harm_Callback(hObject, eventdata, handles)
user_entry = str2double(get(hObject,'string'));
handles.Harm = user_entry;
guidata(hObject,handles)


% --- Executes on button press in PlotDistortion.
function PlotDistortion_Callback(hObject, eventdata, handles)
LineBusXcoord=handles.LineBusXcoord;
LineBusYcoord=handles.LineBusYcoord;
LineBus2Xcoord=handles.LineBus2Xcoord;
LineBus2Ycoord=handles.LineBus2Ycoord;
Harm = handles.Harm;
THDv = zeros(length(LineBusXcoord),1);
Dist = zeros(length(LineBusXcoord),1);
Vh = zeros(length(LineBusXcoord),1);
THDi = zeros(length(LineBusXcoord),1);
iDist = zeros(length(LineBusXcoord),1);
inDist = zeros(length(LineBusXcoord),1);
vUnbal = zeros(length(LineBusXcoord),1);
iUnbal = zeros(length(LineBusXcoord),1);
h = waitbar(0,'Please wait...');
var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_vi_',num2str(1),'.csv'],',',1);
for ii=1:length(LineBusXcoord)
    waitbar((ii / length(LineBusXcoord)),h)
    var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_vi_',num2str(ii),'.csv'],',',1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V1'}));
    [null,indV2]=find(strcmp(var1.colheaders,{' V2'}));
    [null,indV3]=find(strcmp(var1.colheaders,{' V3'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' I1'}));
    [null,indI2]=find(strcmp(var1.colheaders,{' I2'}));
    [null,indI3]=find(strcmp(var1.colheaders,{' I3'}));
    [null,indIN]=find(strcmp(var1.colheaders,{' IN'}));
    if isempty(indV3)==0
        THDv(ii,1)=mean([sqrt(sum((var1.data(2:100,indV1)./var1.data(1,indV1)).^2))*100,...
            sqrt(sum((var1.data(2:100,indV2)./var1.data(1,indV2)).^2))*100,...
            sqrt(sum((var1.data(2:100,indV3)./var1.data(1,indV3)).^2))*100]);
        Dist(ii,1)=mean([var1.data(Harm,indV1)/var1.data(1,indV1)*100,...
            var1.data(Harm,indV2)/var1.data(1,indV2)*100,...
            var1.data(Harm,indV3)/var1.data(1,indV3)*100]);
        Vh(ii,1)=mean([var1.data(Harm,indV1),var1.data(Harm,indV2),var1.data(Harm,indV3)]);
        THDi(ii,1)=mean([sqrt(sum((var1.data(2:100,indI1)./var1.data(1,indI1)).^2))*100,...
            sqrt(sum((var1.data(2:100,indI2)./var1.data(1,indI2)).^2))*100,...
            sqrt(sum((var1.data(2:100,indI3)./var1.data(1,indI3)).^2))*100]);
        iDist(ii,1)=mean([var1.data(Harm,indI1),...
            var1.data(Harm,indI2),...
            var1.data(Harm,indI3)]);
        Vrms1=sqrt(sum((var1.data(1:100,indV1)).^2));
        Vrms2=sqrt(sum((var1.data(1:100,indV2)).^2));
        Vrms3=sqrt(sum((var1.data(1:100,indV3)).^2));
        avgV=mean([Vrms1,Vrms2,Vrms3]);
        vUnbal(ii,1)=max([abs(Vrms1-avgV),abs(Vrms2-avgV),abs(Vrms3-avgV)])/avgV*100;
        Irms1=sqrt(sum((var1.data(1:100,indI1)).^2));
        Irms2=sqrt(sum((var1.data(1:100,indI2)).^2));
        Irms3=sqrt(sum((var1.data(1:100,indI3)).^2));
        avgI=mean([Irms1,Irms2,Irms3]);
        iUnbal(ii,1)=max([abs(Irms1-avgI),abs(Irms2-avgI),abs(Irms3-avgI)])/avgI*100;
    elseif isempty(indV2)==0
        THDv(ii,1)=mean([sqrt(sum((var1.data(2:100,indV1)./var1.data(1,indV1)).^2))*100,...
            sqrt(sum((var1.data(2:100,indV2)./var1.data(1,indV2)).^2))*100]);
        Dist(ii,1)=mean([var1.data(Harm,indV1)/var1.data(1,indV1)*100,...
            var1.data(Harm,indV2)/var1.data(1,indV2)*100]);
        Vh(ii,1)=mean([var1.data(Harm,indV1),var1.data(Harm,indV2)]);
        THDi(ii,1)=mean([sqrt(sum((var1.data(2:100,indI1)./var1.data(1,indI1)).^2))*100,...
            sqrt(sum((var1.data(2:100,indI2)./var1.data(1,indI2)).^2))*100]);
        iDist(ii,1)=mean([var1.data(Harm,indI1),...
            var1.data(Harm,indI2)]);
        Vrms1=sqrt(sum((var1.data(1:100,indV1)).^2));
        Vrms2=sqrt(sum((var1.data(1:100,indV2)).^2));
        avgV=mean([Vrms1,Vrms2]);
        vUnbal(ii,1)=max([abs(Vrms1-avgV),abs(Vrms2-avgV)])/avgV*100;
        Irms1=sqrt(sum((var1.data(1:100,indI1)).^2));
        Irms2=sqrt(sum((var1.data(1:100,indI2)).^2));
        avgI=mean([Irms1,Irms2]);
        iUnbal(ii,1)=max([abs(Irms1-avgI),abs(Irms2-avgI)])/avgI*100;
    else
        THDv(ii,1)=sqrt(sum((var1.data(2:100,indV1)./var1.data(1,indV1)).^2))*100;
        Dist(ii,1)=var1.data(Harm,indV1)/var1.data(1,indV1)*100;
        Vh(ii,1)=var1.data(Harm,indV1);
        THDi(ii,1)=sqrt(sum((var1.data(2:100,indI1)./var1.data(1,indI1)).^2))*100;
        iDist(ii,1)=mean([var1.data(Harm,indI1)]);
    end
    if isempty(indIN)==0,inDist(ii,1)=sqrt(sum((var1.data(1:100,indIN)).^2)); end
end
close(h)

scrzx=get(0,'ScreenSize');
figure(10)
close(10)
figure(10)
jetcolors = jet;
set(10,'outerposition',[0 scrzx(4)-800 800 800],'Name','Voltage')
subplot(2,2,1)
% scatter(LineBusXcoord,LineBusYcoord,[],Dist,'filled')
delta = (max(Dist)-min(Dist));
if delta==0, delta=0.0001; end
for ii=1:length(Dist)
    if Dist(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((Dist(ii)-min(Dist))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(Dist) min(Dist)+delta]);
colormap('jet');
title(['Average Phase Voltage Distortion (%) - Harmonic ',num2str(Harm)],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar
subplot(2,2,2)
% scatter(LineBusXcoord,LineBusYcoord,[],THDv,'filled')
delta = (max(THDv)-min(THDv));
if delta==0, delta=0.0001; end
for ii=1:length(THDv)
    if THDv(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((THDv(ii)-min(THDv))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(THDv) min(THDv)+delta]);
colormap('jet');
title(['Average Phase THDv (%)'],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar
subplot(2,2,3)
% scatter(LineBusXcoord,LineBusYcoord,[],vUnbal,'filled')
delta = (max(vUnbal)-min(vUnbal));
if delta==0, delta=0.0001; end
for ii=1:length(vUnbal)
    if vUnbal(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((vUnbal(ii)-min(vUnbal))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(vUnbal) min(vUnbal)+delta]);
colormap('jet');
title(['Maximum Phase RMS Voltage Unbalance (%)'],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar;
subplot(2,2,4)
% scatter(LineBusXcoord,LineBusYcoord,[],Dist,'filled')
delta = (max(Vh)-min(Vh));
if delta==0, delta=0.0001; end
for ii=1:length(Vh)
    if Vh(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((Vh(ii)-min(Vh))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(Vh) min(Vh)+delta]);
colormap('jet');
title(['Average Phase Voltage (V) - Harmonic ',num2str(Harm)],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar

figure(11)
close(11)
figure(11)
jetcolors = jet;
set(11,'outerposition',[0 scrzx(4)-800 800 800],'Name','Current')
subplot(2,2,1)
% scatter(LineBusXcoord,LineBusYcoord,[],iDist,'filled')
delta = (max(iDist)-min(iDist));
if delta==0, delta=0.0001; end
for ii=1:length(iDist)
    if iDist(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((iDist(ii)-min(iDist))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(iDist) min(iDist)+delta]);
colormap('jet');
title(['Average Phase Current (A) - Harmonic ',num2str(Harm)],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar
subplot(2,2,2)
% scatter(LineBusXcoord,LineBusYcoord,[],THDi,'filled')
delta = (max(THDi)-min(THDi));
if delta==0, delta=0.0001; end
for ii=1:length(THDi)
    if THDi(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((THDi(ii)-min(THDi))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(THDi) min(THDi)+delta]);
colormap('jet');
title([{'Average Phase THDi (%)'}],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar
subplot(2,2,4)
% scatter(LineBusXcoord,LineBusYcoord,[],inDist,'filled')
delta = (max(inDist)-min(inDist));
if delta==0, delta=0.0001; end
for ii=1:length(inDist)
    if inDist(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((inDist(ii)-min(inDist))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(inDist) min(inDist)+delta]);
colormap('jet');
title(['RMS Neutral Current (A)'],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar
subplot(2,2,3)
% scatter(LineBusXcoord,LineBusYcoord,[],iUnbal,'filled')
delta = (max(iUnbal)-min(iUnbal));
if delta==0, delta=0.0001; end
for ii=1:length(iUnbal)
    if iUnbal(ii)>=0, plot([LineBusXcoord(ii),LineBus2Xcoord(ii)],[LineBusYcoord(ii),LineBus2Ycoord(ii)],'linewidth',3,'color',jetcolors(round((iUnbal(ii)-min(iUnbal))/delta*(length(jetcolors)-1))+1,:,:)); hold on; end
end
caxis([min(iUnbal) min(iUnbal)+delta]);
colormap('jet');
title(['Maximum Phase RMS Current Unbalance (%)'],'fontsize',14)
set(gca,'Xtick',[],'Ytick',[])
colorbar

print(['-f10'],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/05_VoltageSchematic_Bus-',handles.DistortionBus,'_Harmonic-',num2str(handles.Harm)])
print(['-f11'],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/06_CurrentSchematic_Bus-',handles.DistortionBus,'_Harmonic-',num2str(handles.Harm)'])





% --- Executes on button press in LoadDistortion.
function LoadDistortion_Callback(hObject, eventdata, handles)
data1=zeros(1,27);
data2=zeros(1,27);
data3=zeros(1,27);
data4=zeros(1,27);
idata1=zeros(1,27);
idata2=zeros(1,27);
idata3=zeros(1,27);
idata4=zeros(1,27);
if handles.AllCapConfigs2==0
    var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_vi.csv'],',',1);
    
    Distortion=zeros(size(var1.data,1)-1,1);
    iDistortion=zeros(size(var1.data,1)-1,1);
    THDiDistortion=zeros(size(var1.data,1)-1,1);
    Harmonic=zeros(size(var1.data,1)-1,1);
    Freq=zeros(size(var1.data,1)-1,1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V1'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' I1'}));
    [null,indHarm]=find(strcmp(var1.colheaders,{' Harmonic'}));
    Vrms = var1.data(1,indV1).^2;
    Irms = var1.data(1,indI1).^2;
    for jj=1:size(var1.data,1)-1
        Distortion(jj)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
        THDiDistortion(jj)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
        iDistortion(jj)=var1.data(jj+1,indI1);
        Harmonic(jj)=var1.data(jj+1,indHarm);
        Vrms = Vrms + var1.data(jj+1,indV1).^2;
        Irms = Irms + var1.data(jj+1,indI1).^2;
    end
    THDv=sqrt(sum(Distortion.^2));
    THDi=sqrt(sum(THDiDistortion.^2));
%     figure(1)
%     scatter(Harmonic,Distortion,[],'r','filled');
%     text(3,0,[{['THDv=',num2str(THDv)]}],'VerticalAlignment','bottom','HorizontalAlignment','left');
%     title(['Bus ',handles.DistortionBus])
%     xlabel('Harmonic')
%     ylabel('Phase A Voltage (% of Fundamental)')
%     hold off;
    data1=[var1.data(1,indV1)/1000,sqrt(Vrms)/1000,THDv,transpose(Distortion(1:99))];
    idata1=[var1.data(1,indI1)/1000,sqrt(Irms)/1000,THDi,transpose(iDistortion(1:99))];
    
    [null,indV1]=find(strcmp(var1.colheaders,{' V2'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' I2'}));
    if isempty(indV1)==0
        Vrms = var1.data(1,indV1).^2;
        Irms = var1.data(1,indI1).^2;
        for jj=1:size(var1.data,1)-1
            Distortion(jj)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
            THDiDistortion(jj)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
            iDistortion(jj)=var1.data(jj+1,indI1);
            Vrms = Vrms + var1.data(jj+1,indV1).^2;
            Irms = Irms + var1.data(jj+1,indI1).^2;
        end
        THDv=sqrt(sum(Distortion.^2));
        THDi=sqrt(sum(THDiDistortion.^2));
%         figure(2)
%         scatter(Harmonic,Distortion,[],'r','filled');
%         text(3,0,[{['THDv=',num2str(THDv)]}],'VerticalAlignment','bottom','HorizontalAlignment','left');
%         title(['Bus ',handles.DistortionBus])
%         xlabel('Harmonic')
%         ylabel('Phase B Voltage (% of Fundamental)')
%         hold off;
        data2=[var1.data(1,indV1)/1000,sqrt(Vrms)/1000,THDv,transpose(Distortion(1:99))];
        idata2=[var1.data(1,indI1)/1000,sqrt(Irms)/1000,THDi,transpose(iDistortion(1:99))];
    end
    [null,indV1]=find(strcmp(var1.colheaders,{' V3'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' I3'}));
    if isempty(indV1)==0
        Vrms = var1.data(1,indV1).^2;
        Irms = var1.data(1,indI1).^2;
        for jj=1:size(var1.data,1)-1
            Distortion(jj)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
            THDiDistortion(jj)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
            iDistortion(jj)=var1.data(jj+1,indI1);
            Vrms = Vrms + var1.data(jj+1,indV1).^2;
            Irms = Irms + var1.data(jj+1,indI1).^2;
        end
        THDv=sqrt(sum(Distortion.^2));
        THDi=sqrt(sum(THDiDistortion.^2));
%         figure(3)
%         scatter(Harmonic,Distortion,[],'r','filled');
%         text(3,0,[{['THDv=',num2str(THDv)]}],'VerticalAlignment','bottom','HorizontalAlignment','left');
%         title(['Bus ',handles.DistortionBus])
%         xlabel('Harmonic')
%         ylabel('Phase C Voltage (% of Fundamental)')
%         hold off;
        data3=[var1.data(1,indV1)/1000,sqrt(Vrms)/1000,THDv,transpose(Distortion(1:99))];
        idata3=[var1.data(1,indI1)/1000,sqrt(Irms)/1000,THDi,transpose(iDistortion(1:99))];
    end
    [null,indV1]=find(strcmp(var1.colheaders,{' VN'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' IN'}));
    Vrms = var1.data(1,indV1).^2;
    Irms = var1.data(1,indI1).^2;
    for jj=1:size(var1.data,1)-1
        Distortion(jj)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
        THDiDistortion(jj)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
        iDistortion(jj)=var1.data(jj+1,indI1);
        Vrms = Vrms + var1.data(jj+1,indV1).^2;
        Irms = Irms + var1.data(jj+1,indI1).^2;
    end
    THDv=sqrt(sum(Distortion.^2));
    THDi=sqrt(sum(THDiDistortion.^2));
%     figure(4)
%     scatter(Harmonic,Distortion,[],'r','filled');
%     title(['Bus ',handles.DistortionBus])
%     xlabel('Harmonic')
%     ylabel('Neutral-Earth Voltage (% of Fundamental)')
%     hold off;
    data4=[var1.data(1,indV1)/1000,sqrt(Vrms)/1000,-1,transpose(Distortion(1:99))];
    idata4=[var1.data(1,indI1)/1000,sqrt(Irms)/1000,-1,transpose(iDistortion(1:99))];
    
elseif handles.AllCapConfigs2==1
    % Table results for all cap configs is average fundamental, maxTHDv,
    % and max individual harmonic distortion. 
    CapNames=handles.CapNamesAnalyzed2;
    TotalCaps=length(CapNames);
    var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_vi.csv'],',',1);
    
    Distortion=zeros(size(var1.data,1)-1,2^TotalCaps);
    iDistortion=zeros(size(var1.data,1)-1,2^TotalCaps);
    THDiDistortion=zeros(size(var1.data,1)-1,2^TotalCaps);
    Harmonic=zeros(size(var1.data,1)-1,1);
    Freq=zeros(size(var1.data,1)-1,1);
    Vrms=zeros(2^TotalCaps,1);
    Irms=zeros(2^TotalCaps,1);
    Fundamental=zeros(2^TotalCaps,1);
    iFundamental=zeros(2^TotalCaps,1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V1'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' I1'}));
    [null,indHarm]=find(strcmp(var1.colheaders,{' Harmonic'}));
    Vrms(1) = var1.data(1,indV1).^2;
    Irms(1) = var1.data(1,indI1).^2;
    for jj=1:size(var1.data,1)-1
        Distortion(jj,1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
        THDiDistortion(jj,1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
        iDistortion(jj,1)=var1.data(jj+1,indI1);
        Irms(1) = Irms(1) + var1.data(jj+1,indI1).^2;
        Harmonic(jj)=var1.data(jj+1,indHarm);
        Vrms(1) = Vrms(1) + var1.data(jj+1,indV1).^2;
    end
    Fundamental(1)=var1.data(1,indV1);
    iFundamental(1)=var1.data(1,indI1);
    for kk=1:2^TotalCaps-1
        var1 = importdata([handles.mydir,'\Export\Distortion_',num2str(kk),'_Mon_vi.csv'],',',1);
        Vrms(kk+1) = var1.data(1,indV1).^2;
        Irms(kk+1) = var1.data(1,indI1).^2;
        for jj=1:size(var1.data,1)-1
            Distortion(jj,kk+1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
            Vrms(kk+1) = Vrms(kk+1) + var1.data(jj+1,indV1).^2;
            THDiDistortion(jj,kk+1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
            iDistortion(jj,kk+1)=var1.data(jj+1,indI1);
            Irms(kk+1) = Irms(kk+1) + var1.data(jj+1,indI1).^2;
        end
        Fundamental(kk+1)=var1.data(1,indV1);
        iFundamental(kk+1)=var1.data(1,indI1);
    end
    MaxDistortion=max(Distortion,[],2);
    iMaxDistortion=max(iDistortion,[],2);
    MinDistortion=min(Distortion,[],2);
    MeanDistortion=mean(Distortion,2);
    MaxFeederTHDv=sqrt(max(sum(Distortion.^2,1)));
    MaxFeederTHDi=sqrt(max(sum(THDiDistortion.^2,1)));
    MaxNoCapsFeederTHDv=sqrt(sum(Distortion(:,1).^2,1));
    figure(1)
    errorbar(Harmonic,MeanDistortion,MeanDistortion-MinDistortion,MaxDistortion-MeanDistortion,'.'), hold on;
    scatter(Harmonic,Distortion(:,1),[],'ro','filled');
    text(3,0,[{['Max THDv w/Caps=',num2str(MaxFeederTHDv)]},{['Max THDv w/NoCaps=',num2str(MaxNoCapsFeederTHDv)]}],'VerticalAlignment','bottom','HorizontalAlignment','left');
    title(['Bus ',handles.DistortionBus])
    legend('Caps','No Caps')
    xlabel('Harmonic')
    ylabel('Phase 1 Voltage (% of Fundamental)')
    hold off;
    data1=[mean(Fundamental)/1000,mean(sqrt(Vrms))/1000,MaxFeederTHDv,transpose(MaxDistortion(1:99))];
    idata1=[mean(iFundamental)/1000,mean(sqrt(Irms))/1000,MaxFeederTHDi,transpose(iMaxDistortion(1:99))];
    
    Config=[find(sqrt(sum(Distortion.^2,1))==MaxFeederTHDv),...
        find(Distortion(1,:)==MaxDistortion(1)),...
        find(Distortion(2,:)==MaxDistortion(2)),...
        find(Distortion(3,:)==MaxDistortion(3)),...
        find(Distortion(4,:)==MaxDistortion(4)),...
        find(Distortion(5,:)==MaxDistortion(5)),...
        find(Distortion(6,:)==MaxDistortion(6)),...
        find(Distortion(7,:)==MaxDistortion(7)),...
        find(Distortion(8,:)==MaxDistortion(8)),...
        find(Distortion(9,:)==MaxDistortion(9)),...
        find(Distortion(10,:)==MaxDistortion(10)),...
        find(Distortion(11,:)==MaxDistortion(11)),...
        find(Distortion(12,:)==MaxDistortion(12))];
    CapConfigStatus=handles.CapConfigStatus2;
    WorseCaseConfig=[CapConfigStatus(:,Config(1)),...
        CapConfigStatus(:,Config(2)),...
        CapConfigStatus(:,Config(3)),...
        CapConfigStatus(:,Config(4)),...
        CapConfigStatus(:,Config(5)),...
        CapConfigStatus(:,Config(6)),...
        CapConfigStatus(:,Config(7)),...
        CapConfigStatus(:,Config(8)),...
        CapConfigStatus(:,Config(9)),...
        CapConfigStatus(:,Config(10)),...
        CapConfigStatus(:,Config(11)),...
        CapConfigStatus(:,Config(12)),...
        CapConfigStatus(:,Config(13))];
    f = figure(31);
    set(f,'Name','Phase 1 Voltage - Worst Cap Status','Position',[200 600 1000 150],'Resize','off');
    cnames = {'THDv','H3','H5','H7','H9','H11','H13','H15','H17','H19','H21','H23','H25','H27','H29','H31','H33','H35','H37','H39','H41','H43','H45','H47','H49','H51','H53','H55','H57','H59','H61','H63','H65','H67','H69','H71','H73','H75','H77','H79','H81','H83','H85','H87','H89','H91','H93','H95','H97','H99'};
    rnames = CapNames;
    uitable('Parent',f,'Data',WorseCaseConfig,'ColumnName',cnames,...
        'RowName',rnames,'Position',[20 20 960 110]);

    
    var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_vi.csv'],',',1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V2'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' I2'}));
    if isnumeric(indV1)
        Vrms(1) = var1.data(1,indV1).^2;
        Irms(1) = var1.data(1,indI1).^2;
        for jj=1:size(var1.data,1)-1
            Distortion(jj,1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
            THDiDistortion(jj,1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
            iDistortion(jj,1)=var1.data(jj+1,indI1);
            Irms(1) = Irms(1) + var1.data(jj+1,indI1).^2;
            Harmonic(jj)=var1.data(jj+1,indHarm);
            Vrms(1) = Vrms(1) + var1.data(jj+1,indV1).^2;
        end
        Fundamental(1)=var1.data(1,indV1);
        iFundamental(1)=var1.data(1,indI1);
        for kk=1:2^TotalCaps-1
            var1 = importdata([handles.mydir,'\Export\Distortion_',num2str(kk),'_Mon_vi.csv'],',',1);
            Vrms(kk+1) = var1.data(1,indV1).^2;
            Irms(kk+1) = var1.data(1,indI1).^2;
            for jj=1:size(var1.data,1)-1
                Distortion(jj,kk+1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
                Vrms(kk+1) = Vrms(kk+1) + var1.data(jj+1,indV1).^2;
                THDiDistortion(jj,kk+1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
                iDistortion(jj,kk+1)=var1.data(jj+1,indI1);
                Irms(kk+1) = Irms(kk+1) + var1.data(jj+1,indI1).^2;
            end
            Fundamental(kk+1)=var1.data(1,indV1);
            iFundamental(kk+1)=var1.data(1,indI1);
        end
        MaxDistortion=max(Distortion,[],2);
        iMaxDistortion=max(iDistortion,[],2);
        MinDistortion=min(Distortion,[],2);
        MeanDistortion=mean(Distortion,2);
        MaxFeederTHDv=sqrt(max(sum(Distortion.^2,1)));
        MaxFeederTHDi=sqrt(max(sum(THDiDistortion.^2,1)));
        MaxNoCapsFeederTHDv=sqrt(sum(Distortion(:,1).^2,1));
        figure(2)
        errorbar(Harmonic,MeanDistortion,MeanDistortion-MinDistortion,MaxDistortion-MeanDistortion,'.'), hold on;
        scatter(Harmonic,Distortion(:,1),[],'ro','filled');
        text(3,0,[{['Max THDv w/Caps=',num2str(MaxFeederTHDv)]},{['Max THDv w/NoCaps=',num2str(MaxNoCapsFeederTHDv)]}],'VerticalAlignment','bottom','HorizontalAlignment','left');
        title(['Bus ',handles.DistortionBus])
        legend('Caps','No Caps')
        xlabel('Harmonic')
        ylabel('Phase 2 Voltage (% of Fundamental)')
        hold off;
        data2=[mean(Fundamental)/1000,mean(sqrt(Vrms))/1000,MaxFeederTHDv,transpose(MaxDistortion(1:99))];
        idata2=[mean(iFundamental)/1000,mean(sqrt(Irms))/1000,MaxFeederTHDi,transpose(iMaxDistortion(1:99))];

        Config=[find(sqrt(sum(Distortion.^2,1))==MaxFeederTHDv),...
            find(Distortion(1,:)==MaxDistortion(1)),...
            find(Distortion(2,:)==MaxDistortion(2)),...
            find(Distortion(3,:)==MaxDistortion(3)),...
            find(Distortion(4,:)==MaxDistortion(4)),...
            find(Distortion(5,:)==MaxDistortion(5)),...
            find(Distortion(6,:)==MaxDistortion(6)),...
            find(Distortion(7,:)==MaxDistortion(7)),...
            find(Distortion(8,:)==MaxDistortion(8)),...
            find(Distortion(9,:)==MaxDistortion(9)),...
            find(Distortion(10,:)==MaxDistortion(10)),...
            find(Distortion(11,:)==MaxDistortion(11)),...
            find(Distortion(12,:)==MaxDistortion(12))];
        CapConfigStatus=handles.CapConfigStatus2;
        WorseCaseConfig=[CapConfigStatus(:,Config(1)),...
            CapConfigStatus(:,Config(2)),...
            CapConfigStatus(:,Config(3)),...
            CapConfigStatus(:,Config(4)),...
            CapConfigStatus(:,Config(5)),...
            CapConfigStatus(:,Config(6)),...
            CapConfigStatus(:,Config(7)),...
            CapConfigStatus(:,Config(8)),...
            CapConfigStatus(:,Config(9)),...
            CapConfigStatus(:,Config(10)),...
            CapConfigStatus(:,Config(11)),...
            CapConfigStatus(:,Config(12)),...
            CapConfigStatus(:,Config(13))];
        f = figure(32);
        set(f,'Name','Phase 2 Voltage - Worst Cap Status','Position',[200 500 1000 150],'Resize','off');
        cnames = {'THDv','H3','H5','H7','H9','H11','H13','H15','H17','H19','H21','H23','H25','H27','H29','H31','H33','H35','H37','H39','H41','H43','H45','H47','H49','H51','H53','H55','H57','H59','H61','H63','H65','H67','H69','H71','H73','H75','H77','H79','H81','H83','H85','H87','H89','H91','H93','H95','H97','H99'};
        rnames = CapNames;
        uitable('Parent',f,'Data',WorseCaseConfig,'ColumnName',cnames,...
            'RowName',rnames,'Position',[20 20 960 110]);
    end
    
    var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_vi.csv'],',',1);
    [null,indV1]=find(strcmp(var1.colheaders,{' V3'}));
    if isnumeric(indV1)
        Vrms(1) = var1.data(1,indV1).^2;
        Irms(1) = var1.data(1,indI1).^2;
        for jj=1:size(var1.data,1)-1
            Distortion(jj,1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
            THDiDistortion(jj,1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
            iDistortion(jj,1)=var1.data(jj+1,indI1);
            Irms(1) = Irms(1) + var1.data(jj+1,indI1).^2;
            Harmonic(jj)=var1.data(jj+1,indHarm);
            Vrms(1) = Vrms(1) + var1.data(jj+1,indV1).^2;
        end
        Fundamental(1)=var1.data(1,indV1);
        iFundamental(1)=var1.data(1,indI1);
        for kk=1:2^TotalCaps-1
            var1 = importdata([handles.mydir,'\Export\Distortion_',num2str(kk),'_Mon_vi.csv'],',',1);
            Vrms(kk+1) = var1.data(1,indV1).^2;
            Irms(kk+1) = var1.data(1,indI1).^2;
            for jj=1:size(var1.data,1)-1
                Distortion(jj,kk+1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
                Vrms(kk+1) = Vrms(kk+1) + var1.data(jj+1,indV1).^2;
                THDiDistortion(jj,kk+1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
                iDistortion(jj,kk+1)=var1.data(jj+1,indI1);
                Irms(kk+1) = Irms(kk+1) + var1.data(jj+1,indI1).^2;
            end
            Fundamental(kk+1)=var1.data(1,indV1);
            iFundamental(kk+1)=var1.data(1,indI1);
        end
        MaxDistortion=max(Distortion,[],2);
        iMaxDistortion=max(iDistortion,[],2);
        MinDistortion=min(Distortion,[],2);
        MeanDistortion=mean(Distortion,2);
        MaxFeederTHDv=sqrt(max(sum(Distortion.^2,1)));
        MaxFeederTHDi=sqrt(max(sum(THDiDistortion.^2,1)));
        MaxNoCapsFeederTHDv=sqrt(sum(Distortion(:,1).^2,1));
        figure(3)
        errorbar(Harmonic,MeanDistortion,MeanDistortion-MinDistortion,MaxDistortion-MeanDistortion,'.'), hold on;
        scatter(Harmonic,Distortion(:,1),[],'ro','filled');
        text(3,0,[{['Max THDv w/Caps=',num2str(MaxFeederTHDv)]},{['Max THDv w/NoCaps=',num2str(MaxNoCapsFeederTHDv)]}],'VerticalAlignment','bottom','HorizontalAlignment','left');
        title(['Bus ',handles.DistortionBus])
        legend('Caps','No Caps')
        xlabel('Harmonic')
        ylabel('Phase 3 Voltage (% of Fundamental)')
        hold off;
        data3=[mean(Fundamental)/1000,mean(sqrt(Vrms))/1000,MaxFeederTHDv,transpose(MaxDistortion(1:99))];
        idata3=[mean(iFundamental)/1000,mean(sqrt(Irms))/1000,MaxFeederTHDi,transpose(iMaxDistortion(1:99))];

        Config=[find(sqrt(sum(Distortion.^2,1))==MaxFeederTHDv),...
            find(Distortion(1,:)==MaxDistortion(1)),...
            find(Distortion(2,:)==MaxDistortion(2)),...
            find(Distortion(3,:)==MaxDistortion(3)),...
            find(Distortion(4,:)==MaxDistortion(4)),...
            find(Distortion(5,:)==MaxDistortion(5)),...
            find(Distortion(6,:)==MaxDistortion(6)),...
            find(Distortion(7,:)==MaxDistortion(7)),...
            find(Distortion(8,:)==MaxDistortion(8)),...
            find(Distortion(9,:)==MaxDistortion(9)),...
            find(Distortion(10,:)==MaxDistortion(10)),...
            find(Distortion(11,:)==MaxDistortion(11)),...
            find(Distortion(12,:)==MaxDistortion(12))];
        CapConfigStatus=handles.CapConfigStatus2;
        WorseCaseConfig=[CapConfigStatus(:,Config(1)),...
            CapConfigStatus(:,Config(2)),...
            CapConfigStatus(:,Config(3)),...
            CapConfigStatus(:,Config(4)),...
            CapConfigStatus(:,Config(5)),...
            CapConfigStatus(:,Config(6)),...
            CapConfigStatus(:,Config(7)),...
            CapConfigStatus(:,Config(8)),...
            CapConfigStatus(:,Config(9)),...
            CapConfigStatus(:,Config(10)),...
            CapConfigStatus(:,Config(11)),...
            CapConfigStatus(:,Config(12)),...
            CapConfigStatus(:,Config(13))];
        f = figure(33);
        set(f,'Name','Phase 3 Voltage - Worst Cap Status','Position',[200 400 1000 150],'Resize','off');
        cnames = {'THDv','H3','H5','H7','H9','H11','H13','H15','H17','H19','H21','H23','H25','H27','H29','H31','H33','H35','H37','H39','H41','H43','H45','H47','H49','H51','H53','H55','H57','H59','H61','H63','H65','H67','H69','H71','H73','H75','H77','H79','H81','H83','H85','H87','H89','H91','H93','H95','H97','H99'};
        rnames = CapNames;
        uitable('Parent',f,'Data',WorseCaseConfig,'ColumnName',cnames,...
            'RowName',rnames,'Position',[20 20 960 110]);
    end
    
    var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_vi.csv'],',',1);
    [null,indV1]=find(strcmp(var1.colheaders,{' VN'}));
    [null,indI1]=find(strcmp(var1.colheaders,{' IN'}));
    Vrms(1) = var1.data(1,indV1).^2;
    Irms(1) = var1.data(1,indI1).^2;
    for jj=1:size(var1.data,1)-1
        Distortion(jj,1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
        Vrms(1) = Vrms(1) + var1.data(jj+1,indV1).^2;
        THDiDistortion(jj,1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
        iDistortion(jj,1)=var1.data(jj+1,indI1);
        Irms(1) = Irms(1) + var1.data(jj+1,indI1).^2;
    end
    Fundamental(1)=var1.data(1,indV1);
    iFundamental(1)=var1.data(1,indI1);
    for kk=1:2^TotalCaps-1
        var1 = importdata([handles.mydir,'\Export\Distortion_',num2str(kk),'_Mon_vi.csv'],',',1);
        Vrms(kk+1) = var1.data(1,indV1).^2;
        Irms(kk+1) = var1.data(1,indI1).^2;
        for jj=1:size(var1.data,1)-1
            Distortion(jj,kk+1)=var1.data(jj+1,indV1)/var1.data(1,indV1)*100;
            Vrms(kk+1) = Vrms(kk+1) + var1.data(jj+1,indV1).^2;
            THDiDistortion(jj,kk+1)=var1.data(jj+1,indI1)/var1.data(1,indI1)*100;
            iDistortion(jj,kk+1)=var1.data(jj+1,indI1);
            Irms(kk+1) = Irms(kk+1) + var1.data(jj+1,indI1).^2;
        end
        Fundamental(kk+1)=var1.data(1,indV1);
        iFundamental(kk+1)=var1.data(1,indI1);
    end
    MaxDistortion=max(Distortion,[],2);
    iMaxDistortion=max(iDistortion,[],2);
    MinDistortion=min(Distortion,[],2);
    MeanDistortion=mean(Distortion,2);
    MaxFeederTHDv=sqrt(max(sum(Distortion.^2,1)));
    MaxFeederTHDi=sqrt(max(sum(THDiDistortion.^2,1)));
    MaxNoCapsFeederTHDv=sqrt(sum(Distortion(:,1).^2,1));
%     figure(4)
%     errorbar(Harmonic,MeanDistortion,MeanDistortion-MinDistortion,MaxDistortion-MeanDistortion,'.'), hold on;
%     scatter(Harmonic,Distortion(:,1),[],'ro','filled');
% %     text(3,0,[{['Max THDv w/Caps=',num2str(MaxFeederTHDv)]},{['Max THDv w/NoCaps=',num2str(MaxNoCapsFeederTHDv)]}],'VerticalAlignment','bottom','HorizontalAlignment','left');
%     title(['Bus ',handles.DistortionBus])
%     legend('Caps','No Caps')
%     xlabel('Harmonic')
%     ylabel('Neutral-Earth Voltage (% of Fundamental)')
%     hold off;
    data4=[mean(Fundamental)/1000,mean(sqrt(Vrms))/1000,-1,transpose(MaxDistortion(1:99))];
    idata4=[mean(iFundamental)/1000,mean(sqrt(Irms))/1000,-1,transpose(iMaxDistortion(1:99))];

end

if data2(1)==0, avgdata = data1; iavgdata = idata1;
elseif data3(1)==0, avgdata = mean([data1;data2],1); iavgdata = mean([idata1;idata2],1);
else avgdata = mean([data1;data2;data3],1); iavgdata = mean([idata1;idata2;idata3],1);
end

f = figure(30);
set(f,'Name','Monitored POI Bus Voltage Distortion','Position',[200 300 1000 200],'Resize','off');
cnames = {'Fundamental (kV)','Vrms (kV)','THDv (%)','H2 (%)','H3 (%)','H4 (%)','H5 (%)','H6 (%)','H7 (%)','H8 (%)','H9 (%)','H10 (%)','H11 (%)','H12 (%)','H13 (%)','H14 (%)','H15 (%)','H16 (%)','H17 (%)','H18 (%)','H19 (%)','H20 (%)','H21 (%)','H22 (%)','H23 (%)','H24 (%)','H25 (%)','H26 (%)','H27 (%)','H28 (%)','H29 (%)','H30 (%)','H31 (%)','H32 (%)','H33 (%)','H34 (%)','H35 (%)','H36 (%)','H37 (%)','H38 (%)','H39 (%)','H40 (%)','H41 (%)','H42 (%)','H43 (%)','H44 (%)','H45 (%)','H46 (%)','H47 (%)','H48 (%)','H49 (%)','H50 (%)','H51 (%)','H52 (%)','H53 (%)','H54 (%)','H55 (%)','H56 (%)','H57 (%)','H58 (%)','H59 (%)','H60 (%)','H61 (%)','H62 (%)','H63 (%)','H64 (%)','H65 (%)','H66 (%)','H67 (%)','H68 (%)','H69 (%)','H70 (%)','H71 (%)','H72 (%)','H73 (%)','H74 (%)','H75 (%)','H76 (%)','H77 (%)','H78 (%)','H79 (%)','H80 (%)','H81 (%)','H82 (%)','H83 (%)','H84 (%)','H85 (%)','H86 (%)','H87 (%)','H88 (%)','H89 (%)','H90 (%)','H91 (%)','H92 (%)','H93 (%)','H94 (%)','H95 (%)','H96 (%)','H97 (%)','H98 (%)','H99 (%)','H100 (%)'};
rnames = {'Phase 1','Phase 2','Phase 3','Phase Average'};%,'Neutral-Earth'};
uitable('Parent',f,'Data',[data1;data2;data3;avgdata],'ColumnName',cnames,...%;data4],'ColumnName',cnames,...
    'RowName',rnames,'Position',[20 20 960 160]);
f = figure(50);
set(f,'Name','Monitored Element Current Distortion','Position',[200 200 1000 200],'Resize','off');
cnames = {'Fundamental (kA)','Irms (kA)','THDi (%)','H2 (A)','H3 (A)','H4 (A)','H5 (A)','H6 (A)','H7 (A)','H8 (A)','H9 (A)','H10 (A)','H11 (A)','H12 (A)','H13 (A)','H14 (A)','H15 (A)','H16 (A)','H17 (A)','H18 (A)','H19 (A)','H20 (A)','H21 (A)','H22 (A)','H23 (A)','H24 (A)','H25 (A)','H26 (A)','H27 (A)','H28 (A)','H29 (A)','H30 (A)','H31 (A)','H32 (A)','H33 (A)','H34 (A)','H35 (A)','H36 (A)','H37 (A)','H38 (A)','H39 (A)','H40 (A)','H41 (A)','H42 (A)','H43 (A)','H44 (A)','H45 (A)','H46 (A)','H47 (A)','H48 (A)','H49 (A)','H50 (A)','H51 (A)','H52 (A)','H53 (A)','H54 (A)','H55 (A)','H56 (A)','H57 (A)','H58 (A)','H59 (A)','H60 (A)','H61 (A)','H62 (A)','H63 (A)','H64 (A)','H65 (A)','H66 (A)','H67 (A)','H68 (A)','H69 (A)','H70 (A)','H71 (A)','H72 (A)','H73 (A)','H74 (A)','H75 (A)','H76 (A)','H77 (A)','H78 (A)','H79 (A)','H80 (A)','H81 (A)','H82 (A)','H83 (A)','H84 (A)','H85 (A)','H86 (A)','H87 (A)','H88 (A)','H89 (A)','H90 (A)','H91 (A)','H92 (A)','H93 (A)','H94 (A)','H95 (A)','H96 (A)','H97 (A)','H98 (A)','H99 (A)','H100 (A)'};
rnames = {'Phase 1','Phase 2','Phase 3','Phase Average'};%,'Neutral-Earth'};
uitable('Parent',f,'Data',[idata1;idata2;idata3;iavgdata],'ColumnName',cnames,...%;idata4],'ColumnName',cnames,...
    'RowName',rnames,'Position',[20 20 960 160]);

%----Voltage Compliance
if avgdata(1)*sqrt(3)<=1
    THDvlimit = 8;
    Hvlimit = 5;
elseif avgdata(1)*sqrt(3)>1 && avgdata(1)*sqrt(3)<=69
    THDvlimit = 5;
    Hvlimit = 3;
elseif avgdata(1)*sqrt(3)>69 && avgdata(1)*sqrt(3)<=161
    THDvlimit = 2.5;
    Hvlimit = 1.5;
else
    THDvlimit = 1.5;
    Hvlimit = 1;
end

Vdata = cell(101,4);
Vdata{2,1} = THDvlimit;
for tt=3:101
    Vdata{tt,1} = Hvlimit;
end
Vdata{1,2} = avgdata(1);
for tt=2:101
    Vdata{tt,2} = avgdata(tt+1);
end

if avgdata(3)>=THDvlimit
    Vdata{2,4} = 'Yes';
    Vdata{2,3} = 'Yes';
elseif avgdata(3)>=THDvlimit*.9
    Vdata{2,3} = 'Yes';
end
for tt=0:49
    if Vdata{tt+2,2}>=Vdata{tt+2,1}
        Vdata{tt+2,4} = 'Yes';
        Vdata{tt+2,3} = 'Yes';
    elseif Vdata{tt+2,2}>Vdata{tt+2,1}*.9
        Vdata{tt+2,3} = 'Yes';
    end
end

f = figure(40);
set(f,'Name','Harmonic Voltage Compliance Reporting','Position',[200 300 600 500],'Resize','off');
rnames = {'POI Voltage (kVln)','THDv (%)','H2 (%)','H3 (%)','H4 (%)','H5 (%)','H6 (%)','H7 (%)','H8 (%)','H9 (%)','H10 (%)','H11 (%)','H12 (%)','H13 (%)','H14 (%)','H15 (%)','H16 (%)','H17 (%)','H18 (%)','H19 (%)','H20 (%)','H21 (%)','H22 (%)','H23 (%)','H24 (%)','H25 (%)','H26 (%)','H27 (%)','H28 (%)','H29 (%)','H30 (%)','H31 (%)','H32 (%)','H33 (%)','H34 (%)','H35 (%)','H36 (%)','H37 (%)','H38 (%)','H39 (%)','H40 (%)','H41 (%)','H42 (%)','H43 (%)','H44 (%)','H45 (%)','H46 (%)','H47 (%)','H48 (%)','H49 (%)','H50 (%)','H51 (%)','H52 (%)','H53 (%)','H54 (%)','H55 (%)','H56 (%)','H57 (%)','H58 (%)','H59 (%)','H60 (%)','H61 (%)','H62 (%)','H63 (%)','H64 (%)','H65 (%)','H66 (%)','H67 (%)','H68 (%)','H69 (%)','H70 (%)','H71 (%)','H72 (%)','H73 (%)','H74 (%)','H75 (%)','H76 (%)','H77 (%)','H78 (%)','H79 (%)','H80 (%)','H81 (%)','H82 (%)','H83 (%)','H84 (%)','H85 (%)','H86 (%)','H87 (%)','H88 (%)','H89 (%)','H90 (%)','H91 (%)','H92 (%)','H93 (%)','H94 (%)','H95 (%)','H96 (%)','H97 (%)','H98 (%)','H99 (%)','H100 (%)'};
cnames = {'Limit','Value','>90% Limit','Violation'};
uitable('Parent',f,'Data',Vdata,'ColumnName',cnames,...
    'RowName',rnames,'Position',[20 20 550 450]);

%----Current Compliance
iscratio = handles.iscDistortionBus/1000/iavgdata(2);
if avgdata(1)*sqrt(3)>161
    if iscratio<25
        TDDilimit = 1.5;
        Hilimit = [0.25,1,0.25,1,0.25,1,0.25,1,0.25,0.5,0.125,0.5,0.125,0.5,0.125,0.38,0.095,0.38,0.095,0.38,0.095,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.1,0.025,0.1,0.025,0.1,0.025,0.1,0.025,0.1,0.025,0.1,0.025,0.1,0.025,0.1,0.025];
    elseif iscratio<50
        TDDilimit = 2.5;
        Hilimit = [0.5,2,0.5,2,0.5,2,0.5,2,0.5,1,0.25,1,0.25,1,0.25,0.75,0.1875,0.75,0.1875,0.75,0.1875,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375];
    else
        TDDilimit = 3.75;
        Hilimit = [0.75,3,0.75,3,0.75,3,0.75,3,0.75,1.5,0.375,1.5,0.375,1.5,0.375,1.15,0.2875,1.15,0.2875,1.15,0.2875,0.45,0.1125,0.45,0.1125,0.45,0.1125,0.45,0.1125,0.45,0.1125,0.45,0.1125,0.22,0.055,0.22,0.055,0.22,0.055,0.22,0.055,0.22,0.055,0.22,0.055,0.22,0.055,0.22,0.055];
    end
elseif avgdata(1)*sqrt(3)>69 && avgdata(1)*sqrt(3)<=161
    if iscratio<20
        TDDilimit = 2.5;
        Hilimit = [0.5,2,0.5,2,0.5,2,0.5,2,0.5,1,0.25,1,0.25,1,0.25,0.75,0.1875,0.75,0.1875,0.75,0.1875,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375,0.15,0.0375];
    elseif iscratio<50
        TDDilimit = 4;
        Hilimit = [0.875,3.5,0.875,3.5,0.875,3.5,0.875,3.5,0.875,1.75,0.4375,1.75,0.4375,1.75,0.4375,1.25,0.3125,1.25,0.3125,1.25,0.3125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.25,0.0625,0.25,0.0625,0.25,0.0625,0.25,0.0625,0.25,0.0625,0.25,0.0625,0.25,0.0625,0.25,0.0625];
    elseif iscratio<100
        TDDilimit = 6;
        Hilimit = [1.25,5,1.25,5,1.25,5,1.25,5,1.25,2.25,0.5625,2.25,0.5625,2.25,0.5625,2,0.5,2,0.5,2,0.5,0.75,0.1875,0.75,0.1875,0.75,0.1875,0.75,0.1875,0.75,0.1875,0.75,0.1875,0.35,0.0875,0.35,0.0875,0.35,0.0875,0.35,0.0875,0.35,0.0875,0.35,0.0875,0.35,0.0875,0.35,0.0875];
    elseif iscratio<1000
        TDDilimit = 7.5;
        Hilimit = [1.5,6,1.5,6,1.5,6,1.5,6,1.5,2.75,0.6875,2.75,0.6875,2.75,0.6875,2.5,0.625,2.5,0.625,2.5,0.625,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125];
    else
        TDDilimit = 10;
        Hilimit = [1.875,7.5,1.875,7.5,1.875,7.5,1.875,7.5,1.875,3.5,0.875,3.5,0.875,3.5,0.875,3,0.75,3,0.75,3,0.75,1.25,0.3125,1.25,0.3125,1.25,0.3125,1.25,0.3125,1.25,0.3125,1.25,0.3125,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175];
    end
elseif avgdata(1)*sqrt(3)<=69
    if iscratio<20
        TDDilimit = 5;
        Hilimit = [1,4,1,4,1,4,1,4,1,2,0.5,2,0.5,2,0.5,1.5,0.375,1.5,0.375,1.5,0.375,0.6,0.15,0.6,0.15,0.6,0.15,0.6,0.15,0.6,0.15,0.6,0.15,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075,0.3,0.075];
    elseif iscratio<50
        TDDilimit = 8;
        Hilimit = [1.75,7,1.75,7,1.75,7,1.75,7,1.75,3.5,0.875,3.5,0.875,3.5,0.875,2.5,0.625,2.5,0.625,2.5,0.625,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125,0.5,0.125];
    elseif iscratio<100
        TDDilimit = 12;
        Hilimit = [2.5,10,2.5,10,2.5,10,2.5,10,2.5,4.5,1.125,4.5,1.125,4.5,1.125,4,1,4,1,4,1,1.5,0.375,1.5,0.375,1.5,0.375,1.5,0.375,1.5,0.375,1.5,0.375,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175,0.7,0.175];
    elseif iscratio<1000
        TDDilimit = 15;
        Hilimit = [3,12,3,12,3,12,3,12,3,5.5,1.375,5.5,1.375,5.5,1.375,5,1.25,5,1.25,5,1.25,2,0.5,2,0.5,2,0.5,2,0.5,2,0.5,2,0.5,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25,1,0.25];
    else
        TDDilimit = 20;
        Hilimit = [3.75,15,3.75,15,3.75,15,3.75,15,3.75,7,1.75,7,1.75,7,1.75,6,1.5,6,1.5,6,1.5,2.5,0.625,2.5,0.625,2.5,0.625,2.5,0.625,2.5,0.625,2.5,0.625,1.4,0.35,1.4,0.35,1.4,0.35,1.4,0.35,1.4,0.35,1.4,0.35,1.4,0.35,1.4,0.35];
    end
end

Idata = cell(53,4);
Idata{4,1} = TDDilimit;
for tt=5:53
    Idata{tt,1} = Hilimit(tt-4);
end
Idata{1,2} = avgdata(1);
Idata{2,2} = handles.iscDistortionBus/1000;
Idata{3,2} = iavgdata(2);
Idata{4,2} = iavgdata(3);
for tt=5:53
    Idata{tt,2} = iavgdata(tt-1)/1000/iavgdata(2)*100;
end

if iavgdata(3)>=TDDilimit
    Idata{4,4} = 'Yes';
    Idata{4,3} = 'Yes';
elseif iavgdata(3)>=TDDilimit*.9
    Idata{4,3} = 'Yes';
end
for tt=0:49
    if Idata{tt+4,2}>=Idata{tt+4,1}
        Idata{tt+4,4} = 'Yes';
        Idata{tt+4,3} = 'Yes';
    elseif Idata{tt+4,2}>Idata{tt+4,1}*.9
        Idata{tt+4,3} = 'Yes';
    end
end

f = figure(41);
set(f,'Name','Harmonic Current Compliance Reporting','Position',[200 200 600 500],'Resize','off');
rnames = {'POI Voltage (kVln)','POI Isc (kA)','POI Iload (kA)','THDi (%)','H2 (%)','H3 (%)','H4 (%)','H5 (%)','H6 (%)','H7 (%)','H8 (%)','H9 (%)','H10 (%)','H11 (%)','H12 (%)','H13 (%)','H14 (%)','H15 (%)','H16 (%)','H17 (%)','H18 (%)','H19 (%)','H20 (%)','H21 (%)','H22 (%)','H23 (%)','H24 (%)','H25 (%)','H26 (%)','H27 (%)','H28 (%)','H29 (%)','H30 (%)','H31 (%)','H32 (%)','H33 (%)','H34 (%)','H35 (%)','H36 (%)','H37 (%)','H38 (%)','H39 (%)','H40 (%)','H41 (%)','H42 (%)','H43 (%)','H44 (%)','H45 (%)','H46 (%)','H47 (%)','H48 (%)','H49 (%)','H50 (%)'};
cnames = {'Limit','Value','>90% Limit','Violation'};
uitable('Parent',f,'Data',Idata,'ColumnName',cnames,...%;idata4],'ColumnName',cnames,...
    'RowName',rnames,'Position',[20 20 550 450]);

var1 = importdata([handles.mydir,'\Export\Distortion_0_Mon_PQ.csv'],',',1);
PQdata = zeros(50,8);
PQdata(:,1:size(var1.data,2)-2) = -var1.data(1:50,3:size(var1.data,2));
PQdata(:,7) = sum([PQdata(:,1),PQdata(:,3),PQdata(:,5)],2);
PQdata(:,8) = sum([PQdata(:,2),PQdata(:,4),PQdata(:,6)],2);
f = figure(42);
set(f,'Name','Monitor Active/Reactive Powers','Position',[200 200 600 500],'Resize','off');
rnames = {'Fundamental','H2','H3','H4','H5','H6','H7','H8','H9','H10','H11','H12','H13','H14','H15','H16','H17','H18','H19','H20','H21','H22','H23','H24','H25','H26','H27','H28','H29','H30','H31','H32','H33','H34','H35','H36','H37','H38','H39','H40','H41','H42','H43','H44','H45','H46','H47','H48','H49','H50'};
cnames = {'P1 (kW)','Q1 (kvar)','P2 (kW)','Q2 (kvar)','P3 (kW)','Q3 (kvar)','Ptotal (kW)','Qtotal (kvar)'};
uitable('Parent',f,'Data',PQdata,'ColumnName',cnames,...%;idata4],'ColumnName',cnames,...
    'RowName',rnames,'Position',[20 20 550 450]);


% print([handles.figure1],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/08'])
print(['-f40'],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/01_VoltageCompliance_Bus-',handles.DistortionBus])
print(['-f41'],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/02_CurrentCompliance_Bus-',handles.DistortionBus])
print(['-f30'],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/03_VoltageDistortion_Bus-',handles.DistortionBus])
print(['-f50'],'-dtiff','-opengl',[handles.mydir,'/Figs/',handles.date,'/04_CurrentDistortion_Bus-',handles.DistortionBus])



% --- Executes during object creation, after setting all properties.
function figure1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes during object deletion, before destroying properties.
function figure1_DeleteFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)




% --- Executes during object creation, after setting all properties.
function ScanType_CreateFcn(hObject, eventdata, handles)
% hObject    handle to ScanType (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function MonitoredElement_CreateFcn(hObject, eventdata, handles)
% hObject    handle to MonitoredElement (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



% --- Executes during object creation, after setting all properties.
function LoadSpectrum_CreateFcn(hObject, eventdata, handles)
% hObject    handle to LoadSpectrum (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes during object creation, after setting all properties.
function SourceSpectrum_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SourceSpectrum (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes during object creation, after setting all properties.
function DGSpectrum_CreateFcn(hObject, eventdata, handles)
% hObject    handle to DGSpectrum (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end





% --- Executes during object creation, after setting all properties.
function loadmult_CreateFcn(hObject, eventdata, handles)
% hObject    handle to loadmult (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end




% --- Executes during object creation, after setting all properties.
function Circuit_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Circuit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end






% --- Executes during object creation, after setting all properties.
function ScanBus_CreateFcn(hObject, eventdata, handles)
% hObject    handle to ScanBus (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end





% --- Executes during object creation, after setting all properties.
function DistortionMonitoredElement_CreateFcn(hObject, eventdata, handles)
% hObject    handle to DistortionMonitoredElement (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function DistortionBus_CreateFcn(hObject, eventdata, handles)
% hObject    handle to DistortionBus (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called





% --- Executes during object creation, after setting all properties.
function Freq_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Freq (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object deletion, before destroying properties.
function CapData_DeleteFcn(hObject, eventdata, handles)
% hObject    handle to CapData (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)




% --- Executes during object creation, after setting all properties.
function Harm_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Harm (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function Isc_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Isc (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes during object creation, after setting all properties.
function Isc1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Isc1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function FSphase_CreateFcn(hObject, eventdata, handles)
% hObject    handle to FSphase (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes during object creation, after setting all properties.
function SeriesRL_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SeriesRL (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on button press in RunFreqScan.
function RunFreqScan_Callback(hObject, eventdata, handles)
set(handles.uipanel4,'Visible','on')
set(handles.uipanel11,'Visible','off')
set(handles.uipanel13,'Visible','off')
set(handles.AboutPanel,'Visible','off')
% set(handles.CktInfo2,'Visible','off')

% --- Executes on button press in CktInfo.
function CktInfo_Callback(hObject, eventdata, handles)
set(handles.uipanel11,'Visible','on')
set(handles.uipanel4,'Visible','off')
set(handles.uipanel13,'Visible','off')
set(handles.AboutPanel,'Visible','off')
% set(handles.CktInfo2,'Visible','off')

% --- Executes on button press in RunDistAnal.
function RunDistAnal_Callback(hObject, eventdata, handles)
set(handles.uipanel13,'Visible','on')
set(handles.uipanel4,'Visible','off')
set(handles.uipanel11,'Visible','off')
set(handles.AboutPanel,'Visible','off')
% set(handles.CktInfo2,'Visible','off')

% --- Executes on button press in About.
function About_Callback(hObject, eventdata, handles)
set(handles.uipanel13,'Visible','off')
set(handles.uipanel4,'Visible','off')
set(handles.uipanel11,'Visible','off')
set(handles.AboutPanel,'Visible','on')
% set(handles.CktInfo2,'Visible','off')


% --- Executes on button press in Accept.
function Accept_Callback(hObject, eventdata, handles)
set(handles.ControlPanel,'Visible','on')
% set(handles.CktInfo2,'Visible','off')
set(handles.SplashScreen,'Visible','off')
set(handles.uipanel11,'Visible','on')
set(handles.uipanel4,'Visible','off')
set(handles.uipanel13,'Visible','off')
set(handles.AboutPanel,'Visible','off')



% --- Executes on button press in Reject.
function Reject_Callback(hObject, eventdata, handles)
close


% --- Executes on button press in Help1.
function Help1_Callback(hObject, eventdata, handles)
%helpdlg('helpstring')
open([cd,'\Help_circuit_selection.pdf'])

% --- Executes on button press in Help2.
function Help2_Callback(hObject, eventdata, handles)
% helpdlg('helpstring')
open([cd,'\Help_FrequencyResponseAnalysis.pdf'])

% --- Executes on button press in Help3.
function Help3_Callback(hObject, eventdata, handles)
% helpdlg('helpstring')
open([cd,'\Help_DistortionAnalysis.pdf'])


% --- Executes on button press in UseNLload.
function UseNLload_Callback(hObject, eventdata, handles)
if get(handles.UseNLload,'Value')==1
    set(handles.UseNLtext,'Visible','on')
else
    set(handles.UseNLtext,'Visible','off')
end


% --- Executes during object creation, after setting all properties.
function edit56_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit56 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end





% --- Executes during object creation, after setting all properties.
function edit57_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit57 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit58_Callback(hObject, eventdata, handles)
% hObject    handle to edit58 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit58 as text
%        str2double(get(hObject,'String')) returns contents of edit58 as a double


% --- Executes during object creation, after setting all properties.
function edit58_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit58 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
