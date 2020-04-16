%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% EMPIRICAL METHODS FOR FINANCE
% Homework III
%
% Benjamin Souane, Antoine-Michel Alexeev and Julien Bisch
% Due Date: 21 April 2020
%==========================================================================

close all
clc
%import KevinShepperd Toolbox
addpath(genpath('C:\Users\41797\Documents\GitHub\EMF_HW3_Group7\mfe-toolbox-master'));

%import functions
addpath(genpath('C:\Users\41797\Documents\GitHub\EMF_HW3_Group7\Functions'))
%%
DataImport;
%% Exporting risk free rate
RF = FFF_5(:,6);
FFF_5 = FFF_5(:,1:5);
%% Exercice 1
% Computing mean excess return for each portfolio
Ex_SBM = mean(Size_BM-RF);
Ex_SOP = mean(Size_OP-RF);
Ex_SINV = mean(Size_INV-RF);

%Computing volatilites of portfolios
Sd_SBM = std(Size_BM);
Sd_SOP = std(Size_OP);
Sd_SINV = std(Size_INV);

% Creating matrices
%Size-BM
name_size = {'Small','size2','size3','size4','Big'};
name_BM = {'BM1','BM2','BM3','BM4','BM5'};
Ex_SBM = reshape(Ex_SBM,[5,5]);
Sd_SBM = reshape(Sd_SBM,[5,5]);
Sharpe_SBM = Ex_SBM./Sd_SBM;

%Size-OP
name_OP = {'OP1','OP2','OP3','OP4','OP5'};
Ex_SOP = reshape(Ex_SOP,[5,5]);
Sd_SOP = reshape(Sd_SOP,[5,5]);
Sharpe_SOP = Ex_SOP./Sd_SOP;

%Size-INV
name_INV = {'INV1','INV2','INV3','INV4','INV5'};
Ex_SINV = reshape(Ex_SINV,[5,5]);
Sd_SINV = reshape(Sd_SINV,[5,5]);
Sharpe_SINV = Ex_SINV./Sd_SINV;

% Creating tables
%SBM
Ex_SBM =array2table(Ex_SBM,'VariableNames',name_size,'RowNames',name_BM);
filename = 'Results/Excess_SBM.xlsx';
writetable(Ex_SBM,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

Sharpe_SBM =array2table(Sharpe_SBM,'VariableNames',name_size,'RowNames',name_BM);
filename = 'Results/Sharpe_SBM.xlsx';
writetable(Sharpe_SBM,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%SOP
Ex_SOP =array2table(Ex_SOP,'VariableNames',name_size,'RowNames',name_OP);
filename = 'Results/Excess_SOP.xlsx';
writetable(Ex_SOP,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

Sharpe_SOP =array2table(Sharpe_SOP,'VariableNames',name_size,'RowNames',name_OP);
filename = 'Results/Sharpe_SBM.xlsx';
writetable(Sharpe_SOP,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%SINV
Ex_SINV =array2table(Ex_SINV,'VariableNames',name_size,'RowNames',name_INV);
filename = 'Results/Excess_SINV.xlsx';
writetable(Ex_SINV,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

Sharpe_SINV =array2table(Sharpe_SINV,'VariableNames',name_size,'RowNames',name_INV);
filename = 'Results/Sharpe_SINV.xlsx';
writetable(Sharpe_SINV,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% Exercice 2
% Computing mean excess return for each portfolio
Mean_FFF5 = mean(FFF_5);
Sd_FFF5 = std(FFF_5);
Sharpe_FFF5 = Mean_FFF5./Sd_FFF5;

%Correlation matrix 
Corr_FFF5 = corrcoef(FFF_5);

% Creating table
M_FFF5 =array2table(Mean_FFF5,'VariableNames',names_FFF5(1:5),'RowNames',{'Excess Return'});
filename = 'Results/Excess_FFF5.xlsx';
writetable(M_FFF5,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

SR_FFF5 =array2table(Sharpe_FFF5,'VariableNames',names_FFF5(1:5),'RowNames',{'Sharpe Ratio'});
filename = 'Results/Sharpe_FFF5.xlsx';
writetable(SR_FFF5,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

Corr_FFF5 =array2table(Corr_FFF5,'VariableNames',names_FFF5(1:5),'RowNames',names_FFF5(1:5));
filename = 'Results/Corr_FFF5.xlsx';
writetable(Corr_FFF5,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% Exercice 3
% a.i
% Computting excess returns
Excess_SBM = Size_BM-RF;

% Setting matrix
[rowSBM,colSBM] = size(Size_BM);
coefficients_SBM = zeros(4,colSBM);
Residuals_SBM = zeros(rowSBM,colSBM);
% Run regressions
for i=1:colSBM
    reg = fitlm(FFF_5(:,1),Excess_SBM(:,i));
    coefficients_SBM(1,i) = reg.Coefficients.Estimate(1);
    coefficients_SBM(3,i) = reg.Coefficients.Estimate(2);
    coefficients_SBM(2,i) = reg.Coefficients.tStat(1);
    coefficients_SBM(4,i) = reg.Coefficients.tStat(2);
    Residuals_SBM(:,i) = reg.Residuals.Raw;
end

coef_SBM = coefficients_SBM(:,[1,5,6,10,11,15,16,20,21,25]);

% Creating table
coef_SBM =array2table(coef_SBM,'VariableNames',{'SizeSmall-LowBM','SizeSmall-HighBM','Size2-LowBM','Size2-HighBM','Size3-LowBM','Size3-HighBM','Size4-LowBM','Size4-HighBM','SizeBig-LowBM','SizeBig-HighBM'},'RowNames',{'alpha','tStat-alpha','beta','tStat-beta'});
filename = 'Results/coef_SBM.xlsx';
writetable(coef_SBM,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% a.ii
% Two stage procedure
%Setting matrix
coefficients_TSP = zeros(2,2);
EM_BM = mean(Size_BM-RF);

% Run regressions
reg = fitlm(coefficients_SBM(3,:)',EM_BM');
coefficients_TSP(1,1) = reg.Coefficients.Estimate(1);
coefficients_TSP(1,2) = reg.Coefficients.Estimate(2);
coefficients_TSP(2,1) = reg.Coefficients.tStat(1);
coefficients_TSP(2,2) = reg.Coefficients.tStat(2);


% Creating table
coefficients_TSP =array2table(coefficients_TSP,'VariableNames',{'Phi0','Phi1'},'RowNames',{'Estimate','t-stat'});
filename = 'Results/coefficients_TSP.xlsx';
writetable(coefficients_TSP,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

% Wald Test
% Covariance matrix
CovMat = 1/rowSBM*(Residuals_SBM'*Residuals_SBM);

% Joint test
J_0 = rowSBM*((1+Sharpe_FFF5(:,1).^2)^(-1))*coefficients_SBM(1,:)*...
    (CovMat^(-1))*coefficients_SBM(1,:)';



% Critical value
CriticalValueWald = chi2inv(0.95, 25);

% Creating table
WaldTest = table([J_0;CriticalValueWald],'VariableNames',{'WaldTest'},'RowNames',{'J_0','Critical Value'});
%% 3.b
%b.i
% Setting matrix
coefficients_3F = zeros(8,colSBM);
Residuals_3F = zeros(rowSBM,colSBM);


% Run regressions
for i=1:colSBM
    reg = fitlm(FFF_5(:,1:3),Excess_SBM(:,i));
    coefficients_3F(1,i) = reg.Coefficients.Estimate(1);
    coefficients_3F(3,i) = reg.Coefficients.Estimate(2);
    coefficients_3F(5,i) = reg.Coefficients.Estimate(3);
    coefficients_3F(7,i) = reg.Coefficients.Estimate(4);
    coefficients_3F(2,i) = reg.Coefficients.tStat(1);
    coefficients_3F(4,i) = reg.Coefficients.tStat(2);
    coefficients_3F(6,i) = reg.Coefficients.tStat(3);
    coefficients_3F(8,i) = reg.Coefficients.tStat(4);
    Residuals_3F(:,i) = reg.Residuals.Raw;
end

coef_3F = coefficients_3F(:,[1,5,6,10,11,15,16,20,21,25]);

% Creating table
coef_3F =array2table(coef_3F,'VariableNames',{'SizeSmall-LowBM','SizeSmall-HighBM','Size2-LowBM','Size2-HighBM','Size3-LowBM','Size3-HighBM','Size4-LowBM','Size4-HighBM','SizeBig-LowBM','SizeBig-HighBM'},'RowNames',{'alpha','tStat-alpha','beta','tStat-beta','s','tStat-s','h','tStat-h'});
filename = 'Results/coef_3F.xlsx';
writetable(coef_3F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% b.ii
% Two stage procedure
%Setting matrix
coefficients_TSP3F = zeros(2,4);

% Run regressions
reg = fitlm(coefficients_3F(3:2:7,:)',EM_BM');
coefficients_TSP3F(1,1) = reg.Coefficients.Estimate(1);
coefficients_TSP3F(1,2) = reg.Coefficients.Estimate(2);
coefficients_TSP3F(1,3) = reg.Coefficients.Estimate(3);
coefficients_TSP3F(1,4) = reg.Coefficients.Estimate(4);
coefficients_TSP3F(2,1) = reg.Coefficients.tStat(1);
coefficients_TSP3F(2,2) = reg.Coefficients.tStat(2);
coefficients_TSP3F(2,3) = reg.Coefficients.tStat(3);
coefficients_TSP3F(2,4) = reg.Coefficients.tStat(4);


% Creating table
coefficients_TSP3F =array2table(coefficients_TSP3F,'VariableNames',{'Phi0','Phi1','Phi2','Phi3'},'RowNames',{'Estimate','t-stat'});
filename = 'Results/coefficients_TSP3F.xlsx';
writetable(coefficients_TSP3F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

% Wald Test

% Covariance matrix
CovMat3F = 1/rowSBM*(Residuals_3F'*Residuals_3F);

%Omega
Omega3F = 1/rowSBM*((FFF_5(:,1:3)-Mean_FFF5(:,1:3))'*(FFF_5(:,1:3)-Mean_FFF5(:,1:3)));

% Joint test
J_03F = rowSBM*((1+Mean_FFF5(:,1:3)*(Omega3F^(-1))*Mean_FFF5(:,1:3)')^(-1))*coefficients_3F(1,:)*...
    (CovMat3F^(-1))*coefficients_3F(1,:)';

J_13F = ((rowSBM-25-3)/25)*((1+Mean_FFF5(:,1:3)*(Omega3F^(-1))*Mean_FFF5(:,1:3)')^(-1))*coefficients_3F(1,:)*...
    (CovMat3F^(-1))*coefficients_3F(1,:)';

%Critical Value F-statistic
CriticalValue3F = finv(0.95,25,rowSBM-25-3);
% Creating table
WaldTest3F = table([J_03F;CriticalValueWald],[J_13F;CriticalValue3F],'VariableNames',{'WaldTest','F-test'},'RowNames',{'J','Critical Value'});

%% 3.c
%c.i
% Setting matrix
coefficients_5F = zeros(12,colSBM);
Residuals_5F = zeros(rowSBM,colSBM);

% Run regressions
for i=1:colSBM
    reg = fitlm(FFF_5,Excess_SBM(:,i));
    coefficients_5F(1,i) = reg.Coefficients.Estimate(1);
    coefficients_5F(3,i) = reg.Coefficients.Estimate(2);
    coefficients_5F(5,i) = reg.Coefficients.Estimate(3);
    coefficients_5F(7,i) = reg.Coefficients.Estimate(4);
    coefficients_5F(9,i) = reg.Coefficients.Estimate(5);
    coefficients_5F(11,i) = reg.Coefficients.Estimate(6);
    coefficients_5F(2,i) = reg.Coefficients.tStat(1);
    coefficients_5F(4,i) = reg.Coefficients.tStat(2);
    coefficients_5F(6,i) = reg.Coefficients.tStat(3);
    coefficients_5F(8,i) = reg.Coefficients.tStat(4);
    coefficients_5F(10,i) = reg.Coefficients.tStat(5);
    coefficients_5F(12,i) = reg.Coefficients.tStat(6);
    Residuals_5F(:,i) = reg.Residuals.Raw;
end

coef_5F = coefficients_5F(:,[1,5,6,10,11,15,16,20,21,25]);

% Creating table
coef_5F =array2table(coef_5F,'VariableNames',{'SizeSmall-LowBM','SizeSmall-HighBM','Size2-LowBM','Size2-HighBM','Size3-LowBM','Size3-HighBM','Size4-LowBM','Size4-HighBM','SizeBig-LowBM','SizeBig-HighBM'},'RowNames',{'alpha','tStat-alpha','beta','tStat-beta','s','tStat-s','h','tStat-h','r','tStat-r','c','tStat-c'});
filename = 'Results/coef_5F.xlsx';
writetable(coef_5F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% c.ii
% Two stage procedure
%Setting matrix
coefficients_TSP5F = zeros(2,6);

% Run regressions
reg = fitlm(coefficients_5F(3:2:11,:)',EM_BM');
coefficients_TSP5F(1,1) = reg.Coefficients.Estimate(1);
coefficients_TSP5F(1,2) = reg.Coefficients.Estimate(2);
coefficients_TSP5F(1,3) = reg.Coefficients.Estimate(3);
coefficients_TSP5F(1,4) = reg.Coefficients.Estimate(4);
coefficients_TSP5F(1,5) = reg.Coefficients.Estimate(5);
coefficients_TSP5F(1,6) = reg.Coefficients.Estimate(6);
coefficients_TSP5F(2,1) = reg.Coefficients.tStat(1);
coefficients_TSP5F(2,2) = reg.Coefficients.tStat(2);
coefficients_TSP5F(2,3) = reg.Coefficients.tStat(3);
coefficients_TSP5F(2,4) = reg.Coefficients.tStat(4);
coefficients_TSP5F(2,5) = reg.Coefficients.tStat(5);
coefficients_TSP5F(2,6) = reg.Coefficients.tStat(6);


% Creating table
coefficients_TSP5F =array2table(coefficients_TSP5F,'VariableNames',{'Phi0','Phi1','Phi2','Phi3','Phi4','Phi5'},'RowNames',{'Estimate','t-stat'});
filename = 'Results/coefficients_TSP5F.xlsx';
writetable(coefficients_TSP5F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

% Wald Test

% Covariance matrix
CovMat5F = 1/rowSBM*(Residuals_5F'*Residuals_5F);

%Omega
Omega5F = 1/rowSBM*((FFF_5-Mean_FFF5)'*(FFF_5-Mean_FFF5));

% Joint test
J_05F = rowSBM*((1+Mean_FFF5*(Omega5F^(-1))*Mean_FFF5')^(-1))*coefficients_5F(1,:)*...
    (CovMat5F^(-1))*coefficients_5F(1,:)';

J_15F = ((rowSBM-25-5)/25)*((1+Mean_FFF5*(Omega5F^(-1))*Mean_FFF5')^(-1))*coefficients_5F(1,:)*...
    (CovMat5F^(-1))*coefficients_5F(1,:)';

%Critical Value F-statistic
CriticalValue5F = finv(0.95,25,rowSBM-25-5);
% Creating table
WaldTest5F = table([J_05F;CriticalValueWald],[J_15F;CriticalValue5F],'VariableNames',{'WaldTest','F-test'},'RowNames',{'J','Critical Value'});

%% Exercice 4
% a.i
% Computting excess returns
Excess_OP = Size_OP-RF;

% Setting matrix
[rowOP,colOP] = size(Size_BM);
coefficients_OP = zeros(4,colOP);
Residuals_OP = zeros(rowOP,colOP);
% Run regressions
for i=1:colOP
    reg = fitlm(FFF_5(:,1),Excess_OP(:,i));
    coefficients_OP(1,i) = reg.Coefficients.Estimate(1);
    coefficients_OP(3,i) = reg.Coefficients.Estimate(2);
    coefficients_OP(2,i) = reg.Coefficients.tStat(1);
    coefficients_OP(4,i) = reg.Coefficients.tStat(2);
    Residuals_OP(:,i) = reg.Residuals.Raw;
end

coef_OP = coefficients_OP(:,[1,5,6,10,11,15,16,20,21,25]);

% Creating table
coef_OP =array2table(coef_OP,'VariableNames',{'SizeSmall-LowOP','SizeSmall-HighOP','Size2-LowOP','Size2-HighOP','Size3-LowOP','Size3-HighOP','Size4-LowOP','Size4-HighOP','SizeBig-LowOP','SizeBig-HighOP'},'RowNames',{'alpha','tStat-alpha','beta','tStat-beta'});
filename = 'Results/coef_OP.xlsx';
writetable(coef_OP,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% a.ii
% Two stage procedure
%Setting matrix
coefficientsOP_TSP = zeros(2,2);
EM_OP = mean(Size_OP-RF);

% Run regressions
reg = fitlm(coefficients_OP(3,:)',EM_OP');
coefficientsOP_TSP(1,1) = reg.Coefficients.Estimate(1);
coefficientsOP_TSP(1,2) = reg.Coefficients.Estimate(2);
coefficientsOP_TSP(2,1) = reg.Coefficients.tStat(1);
coefficientsOP_TSP(2,2) = reg.Coefficients.tStat(2);


% Creating table
coefficientsOP_TSP =array2table(coefficientsOP_TSP,'VariableNames',{'Phi0','Phi1'},'RowNames',{'Estimate','t-stat'});
filename = 'Results/coefficientsOP_TSP.xlsx';
writetable(coefficientsOP_TSP,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

% Wald Test
% Covariance matrix
CovMatOP = 1/rowOP*(Residuals_OP'*Residuals_OP);

% Joint test
J_0_OP = rowOP*((1+Sharpe_FFF5(:,1)^2)^(-1))*coefficients_OP(1,:)*...
    (CovMatOP^(-1))*coefficients_OP(1,:)';

% Creating table
WaldTestOP = table([J_0_OP;CriticalValueWald],'VariableNames',{'WaldTest'},'RowNames',{'J_0','Critical Value'});
%% 3.b
%b.i
% Setting matrix
coefficientsOP_3F = zeros(8,colOP);
ResidualsOP_3F = zeros(rowOP,colOP);

% Run regressions
for i=1:colOP
    reg = fitlm(FFF_5(:,1:3),Excess_OP(:,i));
    coefficientsOP_3F(1,i) = reg.Coefficients.Estimate(1);
    coefficientsOP_3F(3,i) = reg.Coefficients.Estimate(2);
    coefficientsOP_3F(5,i) = reg.Coefficients.Estimate(3);
    coefficientsOP_3F(7,i) = reg.Coefficients.Estimate(4);
    coefficientsOP_3F(2,i) = reg.Coefficients.tStat(1);
    coefficientsOP_3F(4,i) = reg.Coefficients.tStat(2);
    coefficientsOP_3F(6,i) = reg.Coefficients.tStat(3);
    coefficientsOP_3F(8,i) = reg.Coefficients.tStat(4);
    ResidualsOP_3F(:,i) = reg.Residuals.Raw;
end

coefOP_3F = coefficientsOP_3F(:,[1,5,6,10,11,15,16,20,21,25]);

% Creating table
coefOP_3F =array2table(coefOP_3F,'VariableNames',...
        {'SizeSmall-LowBM','SizeSmall-HighOP','Size2-LowOP','Size2-HighOP',...
        'Size3-LowOP','Size3-HighOP','Size4-LowOP','Size4-HighOP','SizeBig-LowOP',...
        'SizeBig-HighOP'},'RowNames',{'alpha','tStat-alpha',...
        'beta','tStat-beta','s','tStat-s','h','tStat-h'});
filename = 'Results/coefOP_3F.xlsx';
writetable(coefOP_3F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% b.ii
% Two stage procedure
%Setting matrix
coefficientsOP_TSP3F = zeros(2,4);

% Run regressions
reg = fitlm(coefficientsOP_3F(3:2:7,:)',EM_OP');
coefficientsOP_TSP3F(1,1) = reg.Coefficients.Estimate(1);
coefficientsOP_TSP3F(1,2) = reg.Coefficients.Estimate(2);
coefficientsOP_TSP3F(1,3) = reg.Coefficients.Estimate(3);
coefficientsOP_TSP3F(1,4) = reg.Coefficients.Estimate(4);
coefficientsOP_TSP3F(2,1) = reg.Coefficients.tStat(1);
coefficientsOP_TSP3F(2,2) = reg.Coefficients.tStat(2);
coefficientsOP_TSP3F(2,3) = reg.Coefficients.tStat(3);
coefficientsOP_TSP3F(2,4) = reg.Coefficients.tStat(4);


% Creating table
coefficientsOP_TSP3F =array2table(coefficientsOP_TSP3F,'VariableNames',{'Phi0','Phi1','Phi2','Phi3'},'RowNames',{'Estimate','t-stat'});
filename = 'Results/coefficientsOP_TSP3F.xlsx';
writetable(coefficientsOP_TSP3F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

% Wald Test

% Covariance matrix
CovMatOP3F = 1/rowOP*(ResidualsOP_3F'*ResidualsOP_3F);

%Omega
OmegaOP3F = 1/rowOP*((FFF_5(:,1:3)-Mean_FFF5(:,1:3))'*(FFF_5(:,1:3)-Mean_FFF5(:,1:3)));

% Joint test
J_0OP3F = rowOP*((1+Mean_FFF5(:,1:3)*(OmegaOP3F^(-1))*Mean_FFF5(:,1:3)')^(-1))*coefficientsOP_3F(1,:)*...
    (CovMatOP3F^(-1))*coefficientsOP_3F(1,:)';

%F-test
J_1OP3F = ((rowOP-25-3)/25)*((1+Mean_FFF5(:,1:3)*(OmegaOP3F^(-1))*Mean_FFF5(:,1:3)')^(-1))*coefficientsOP_3F(1,:)*...
    (CovMatOP3F^(-1))*coefficientsOP_3F(1,:)';

% Creating table
WaldTestOP3F = table([J_0OP3F;CriticalValueWald],[J_1OP3F;CriticalValue3F],'VariableNames',{'WaldTest','F-test'},'RowNames',{'J','Critical Value'});

%% 3.c
%c.i
% Setting matrix
coefficientsOP_5F = zeros(12,colOP);
ResidualsOP_5F = zeros(rowOP,colOP);

% Run regressions
for i=1:colOP
    reg = fitlm(FFF_5,Excess_OP(:,i));
    coefficientsOP_5F(1,i) = reg.Coefficients.Estimate(1);
    coefficientsOP_5F(3,i) = reg.Coefficients.Estimate(2);
    coefficientsOP_5F(5,i) = reg.Coefficients.Estimate(3);
    coefficientsOP_5F(7,i) = reg.Coefficients.Estimate(4);
    coefficientsOP_5F(9,i) = reg.Coefficients.Estimate(5);
    coefficientsOP_5F(11,i) = reg.Coefficients.Estimate(6);
    coefficientsOP_5F(2,i) = reg.Coefficients.tStat(1);
    coefficientsOP_5F(4,i) = reg.Coefficients.tStat(2);
    coefficientsOP_5F(6,i) = reg.Coefficients.tStat(3);
    coefficientsOP_5F(8,i) = reg.Coefficients.tStat(4);
    coefficientsOP_5F(10,i) = reg.Coefficients.tStat(5);
    coefficientsOP_5F(12,i) = reg.Coefficients.tStat(6);
    ResidualsOP_5F(:,i) = reg.Residuals.Raw;
end

coefOP_5F = coefficientsOP_5F(:,[1,5,6,10,11,15,16,20,21,25]);

% Creating table
coefOP_5F =array2table(coefOP_5F,'VariableNames',{'SizeSmall-LowOP','SizeSmall-HighOP',...
    'Size2-LowOP','Size2-HighOP','Size3-LowOP','Size3-HighOP','Size4-LowOP',...
    'Size4-HighOP','SizeBig-LowOP','SizeBig-HighOP'},'RowNames',{'alpha','tStat-alpha','beta',...
            'tStat-beta','s','tStat-s','h','tStat-h','r','tStat-r','c','tStat-c'});
filename = 'Results/coefOP_5F.xlsx';
writetable(coefOP_5F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

%% c.ii
% Two stage procedure
%Setting matrix
coefficientsOP_TSP5F = zeros(2,6);

% Run regressions
reg = fitlm(coefficientsOP_5F(3:2:11,:)',EM_OP');
coefficientsOP_TSP5F(1,1) = reg.Coefficients.Estimate(1);
coefficientsOP_TSP5F(1,2) = reg.Coefficients.Estimate(2);
coefficientsOP_TSP5F(1,3) = reg.Coefficients.Estimate(3);
coefficientsOP_TSP5F(1,4) = reg.Coefficients.Estimate(4);
coefficientsOP_TSP5F(1,5) = reg.Coefficients.Estimate(5);
coefficientsOP_TSP5F(1,6) = reg.Coefficients.Estimate(6);
coefficientsOP_TSP5F(2,1) = reg.Coefficients.tStat(1);
coefficientsOP_TSP5F(2,2) = reg.Coefficients.tStat(2);
coefficientsOP_TSP5F(2,3) = reg.Coefficients.tStat(3);
coefficientsOP_TSP5F(2,4) = reg.Coefficients.tStat(4);
coefficientsOP_TSP5F(2,5) = reg.Coefficients.tStat(5);
coefficientsOP_TSP5F(2,6) = reg.Coefficients.tStat(6);


% Creating table
coefficientsOP_TSP5F =array2table(coefficientsOP_TSP5F,'VariableNames',{'Phi0','Phi1','Phi2','Phi3','Phi4','Phi5'},'RowNames',{'Estimate','t-stat'});
filename = 'Results/coefficientsOP_TSP5F.xlsx';
writetable(coefficientsOP_TSP5F,filename,'Sheet',1,'Range','D1','WriteRowNames',true)

% Wald Test

% Covariance matrix
CovMat5FOP = 1/rowOP*(ResidualsOP_5F'*ResidualsOP_5F);

%Omega
Omega5FOP = 1/rowOP*((FFF_5-Mean_FFF5)'*(FFF_5-Mean_FFF5));

% Joint test
J_0OP5F = rowOP*((1+Mean_FFF5*(Omega5FOP^(-1))*Mean_FFF5')^(-1))*coefficientsOP_5F(1,:)*...
    (CovMat5FOP^(-1))*coefficientsOP_5F(1,:)';

%F-test
J_1OP5F = ((rowOP-25-5)/25)*((1+Mean_FFF5*(Omega5FOP^(-1))*Mean_FFF5')^(-1))*coefficientsOP_5F(1,:)*...
    (CovMat5FOP^(-1))*coefficientsOP_5F(1,:)';

% Creating table
WaldTestOP5F = table([J_0OP5F;CriticalValueWald],[J_1OP5F;CriticalValue5F],'VariableNames',{'WaldTest','F-test'},'RowNames',{'J','Critical Value'});


%% LatexTable
tabletolatex;