%% Setup the Import Options and import the data Size-BM
opts = spreadsheetImportOptions("NumVariables", 26);

% Specify sheet and range
opts.Sheet = "Size-BM";
opts.DataRange = "A16:Z682";

% Specify column names and types
opts.VariableNames = ["date", "SMALL-LoBM", "ME1-BM2", "ME1-BM3", "ME1-BM4", "SMALL-HiBM",...
    "ME2-BM1", "ME2-BM2", "ME2-BM3", "ME2-BM4",	"ME2-BM5", "ME3-BM1",...
    "ME3-BM2", "ME3-BM3", "ME3-BM4", "ME3-BM5", "ME4-BM1", "ME4-BM2",...	
    "ME4-BM3", "ME4-BM4", "ME4-BM5", "BIG-LoBM", "ME5-BM2", "ME5-BM3", "ME5-BM4", "BIG-HiBM"];

opts.VariableTypes = ["string", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% EMPIRICAL METHODS FOR FINANCE
% Homework III
%
% Benjamin Souane, Antoine-Michel Alexeev and Julien Bisch
% Due Date: 21 April 2020
%==========================================================================
% Import the data
Size_BM = readtable("C:\Users\41797\Documents\GitHub\EMF_HW3_Group7\DATA_HW3.xlsx", opts, "UseExcel", false);
clear opts

% Setting date 
date = table2array(Size_BM(:,1));
date = datetime(date,'InputFormat','yyyyMM');

% Creating the vectors we will use
txt = Size_BM.Properties.VariableNames; %Extract the Variables Names
names_SBM = txt(1:end); %Vector of Names (Mainly used for plots)
Size_BM = table2array(Size_BM(2:end,2:end)); %Take out the date from the matrix of price

%% Setup the Import Options and import the data Size-BM
opts = spreadsheetImportOptions("NumVariables", 25);

% Specify sheet and range
opts.Sheet = "Size-OP";
opts.DataRange = "B17:Z683";

% Specify column names and types
opts.VariableNames = ["SMALL-LoOP", "ME1-OP2", "ME1-OP3", "ME1-OP4", "SMALL-HiOP",...
    "ME2-OP1", "ME2-OP2", "ME2-OP3", "ME2-OP4",	"ME2-OP5", "ME3-OP1",...
    "ME3-OP2", "ME3-OP3", "ME3-OP4", "ME3-OP5", "ME4-OP1", "ME4-OP2",...	
    "ME4-OP3", "ME4-OP4", "ME4-OP5", "BIG-LoOP", "ME5-OP2", "ME5-OP3", "ME5-OP4", "BIG-HiOP"];

opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% Import the data
Size_OP = readtable("C:\Users\41797\Documents\GitHub\EMF_HW3_Group7\DATA_HW3.xlsx", opts, "UseExcel", false);
clear opts

% Creating the vectors we will use
txt2 = Size_OP.Properties.VariableNames; %Extract the Variables Names
names_SOP = txt2(1:end); %Vector of Names (Mainly used for plots)
Size_OP = table2array(Size_OP(2:end,1:end)); %Take out the date from the matrix of price

%% Setup the Import Options and import the data Size-BM
opts = spreadsheetImportOptions("NumVariables", 25);

% Specify sheet and range
opts.Sheet = "Size-INV";
opts.DataRange = "B17:Z683";

% Specify column names and types
opts.VariableNames = ["SMALL-LoINV", "ME1-INV2", "ME1-INV3", "ME1-INV4", "SMALL-HiINV",...
    "ME2-INV1", "ME2-INV2", "ME2-INV3", "ME2-INV4",	"ME2-INV5", "ME3-INV1",...
    "ME3-INV2", "ME3-INV3", "ME3-INV4", "ME3-INV5", "ME4-INV1", "ME4-INV2",...	
    "ME4-INV3", "ME4-INV4", "ME4-INV5", "BIG-LoINV", "ME5-INV2", "ME5-INV3", "ME5-INV4", "BIG-HiINV"];

opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% Import the data
Size_INV = readtable("C:\Users\41797\Documents\GitHub\EMF_HW3_Group7\DATA_HW3.xlsx", opts, "UseExcel", false);
clear opts

% Creating the vectors we will use
txt3 = Size_INV.Properties.VariableNames; %Extract the Variables Names
names_SINV = txt3(1:end); %Vector of Names (Mainly used for plots)
Size_INV = table2array(Size_INV(2:end,1:end)); %Take out the date from the matrix of price

%% Setup the Import Options and import the data 5FFF
opts = spreadsheetImportOptions("NumVariables", 6);

% Specify sheet and range
opts.Sheet = "5FFF";
opts.DataRange = "B4:G670";

% Specify column names and types
opts.VariableNames = ["Mkt-RF","SMB","HML","RMW","CMA","RF"];

opts.VariableTypes = ["double", "double", "double", "double", "double", "double"];

% Import the data
FFF_5 = readtable("C:\Users\41797\Documents\GitHub\EMF_HW3_Group7\DATA_HW3.xlsx", opts, "UseExcel", false);
clear opts

% Creating the vectors we will use
txt4 = FFF_5.Properties.VariableNames; %Extract the Variables Names
names_FFF5 = txt4(1:end); %Vector of Names (Mainly used for plots)
FFF_5 = table2array(FFF_5(2:end,1:end)); %Take out the date from the matrix of price