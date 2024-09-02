clear all 
close all 
clc

%% DATA
V_n = 12;
I_n = 10;

R_a = 0.5;
L_a = 0.9e-3;

J=7e-6;
beta=90e-6;

k_v = 0.02;
k_t = 0.02;

%% Transfer Function
tau_a = L_a/R_a;
num_a = 1/R_a;
mu_a=num_a;
den_a= [tau_a 1];
G_a= tf(num_a, den_a);

tau_m=J/beta;
num_m=1/beta;
mu_m=num_m;
den_m=[tau_m 1];
G_m=tf(num_m, den_m);

%% Controllers
T_aa= 0.01;
bandwith_a= 5/T_aa;
k_ia = bandwith_a/mu_a;
k_pa= k_ia*tau_a;

T_am = 100*T_aa;
bandwith_m= 5/T_am;
k_im = bandwith_m/mu_m;
k_pm= k_im*tau_m;


%% Input 

steptime = 2;
stepvalue = 20;
T_r=5e-4;

%% Simulation

% multiple scopes

open_system('example_simulink_scheme2023b');
sim('example_simulink_scheme2023b.slx');
model = 'example_simulink_scheme2023b';



scopeBlocks = find_system(model, 'BlockType', 'Scope');
singlescope = scopeBlocks{1};
allFigures = findall(0, 'Type', 'Figure');
clear allFigures;
allFigures = findall(0, 'Type', 'Figure');
for i=1:length(scopeBlocks)
    open_system(scopeBlocks{i});
    % Retrieve the handle of the Scope block
    scopeHandle = get_param(scopeBlocks{i}, 'Handle');
    % Retrieve the handle of the subsystem
    scopesubsys = get_param(scopeBlocks{i}, 'Parent');
    
    % get all figure handles
    figureHandles = findall(0, 'Type', 'figure', 'Name', get_param(scopeHandle, 'Name'));
    figureHandle=figureHandles(1);
    
    % Save the scope figure as an image    
    scopename = get_param(scopeHandle, 'Name');
    firstSlashPos = find(scopesubsys== '/', 1);
    length_string=strlength(scopesubsys);
    if ~isempty(firstSlashPos)
        resultStr = scopesubsys(firstSlashPos+1:length_string);
    else
        resultStr = '';  % Handle case where there is no '/'
    end


    parts = strsplit(resultStr, '/');
    firstpart=strjoin(parts,'_');
    fileName = strcat(firstpart,'_',scopename,'.emf');


    saveas(figureHandle,fileName);
end
    

close_system(model, 0); % Close the system without saving changes


    
%% CLOSING ALL FIGURES
clc
close all

%% CLOSE ALL SCOPES
shh = get(0,'ShowHiddenHandles');
set(0,'ShowHiddenHandles','On');
hscope = findobj(0,'Type','Figure','Tag','SIMULINK_SIMSCOPE_FIGURE');
close(hscope);
set(0,'ShowHiddenHandles',shh);
