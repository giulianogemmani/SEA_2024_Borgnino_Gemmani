% % single scope example
% 
% open_system('experiment');
% sim('experiment');
% % Find all figure handles
% figHandles = findall(0, 'Type', 'figure');
% 
% % Iterate through handles to identify the scope figure
% for i = 1:length(figHandles)
%     tag = get(figHandles(i), 'Tag');
%     % Check if this is the scope figure
%     if strcmp(tag, 'SIMULINK_SIMSCOPE_FIGURE')
%         % Save the scope figure as an image
%         saveas(figHandles(i), 'scope_output.png');
%         break;
%     end
% end

% multiple scopes example

open_system('PPC_scheme_04072024_a');
sim('PPC_scheme_04072024_a');
model = 'PPC_scheme_04072024_a';
scopeBlocks = find_system(model, 'BlockType', 'Scope');
singlescope = scopeBlocks{1};
allFigures = findall(0, 'Type', 'Figure');
clear allFigures;
allFigures = findall(0, 'Type', 'Figure');
for i=1:length(scopeBlocks)
    open_system(scopeBlocks{i});
    % Retrieve the handle of the Scope block
    scopeHandle = get_param(scopeBlocks{i}, 'Handle');
    scopesubsys = get_param(scopeBlocks{i}, 'Parent');
    % Convert the Scope block to a figure
    %allFigures = findall(0, 'Type', 'Figure');
    figureHandle = findall(0, 'Type', 'figure', 'Name', get_param(scopeHandle, 'Name'));
    %figureHandle = [];
% for i = 1:length(allFigures)
%     fig = allFigures(i);
%     % Check if the figure's UserData or other property matches the Scope block's handle
%     parent = fig.UserData.Parent;
%     if isstruct(fig.UserData) && isfield(fig.UserData, 'BlockHandle') && fig.UserData.BlockHandle == scopeHandle
%         figureHandle = fig;
%         break;
%     end
% end
    % Save the scope figure as an image
    figure = figureHandle(1);    
    figprops = get(figure);
    par = figure.Parent;
    parparam = get(par);
    par2 = par.Parent;
    par2param = get(par2);
    %parent = get_param(figure, 'Parent');
    %figureHandle = get_param(scopeHandle, 'ScopeFigure');
    scopename = get_param(scopeHandle, 'Name');
    slashposition=strfind(scopesubsys,'/');
    length_string=strlength(scopesubsys);
    extractedPortion = scopesubsys((slashposition+1):length_string);
    %fileName=strcat(scopename,'.png')
    fileName = strcat(scopename,'_',extractedPortion,'.png');
    saveas(figure,fileName);
end
    
% for i = 1:length(scopeBlocks)
%     set_param([model, '/', scopeBlocks{i}], 'Open', 'on');
% end

% Find all figure handles

%figHandles = findall(0, 'Type', 'figure');

% Iterate through handles to identify the scope figure
% for i = 1:length(figHandles)
%     tag = get(figHandles(i), 'Tag');
%     % Check if this is the scope figure
%     if strcmp(tag, 'SIMULINK_SIMSCOPE_FIGURE')
%         % Save the scope figure as an image
%         saveas(figHandles(i), 'scope_output_multiple.png');
%         %break;
%     end
% end

% % code to be used with a single scope
% model = 'experiment';
% open_system(model);
% set_param(model, 'SimulationCommand', 'start');
% set_param('experiment/Scope5', 'Open', 'on');
% scopeFigHandle = findobj(0, 'Tag', 'SIMULINK_SIMSCOPE_FIGURE');
% if ~isempty(scopeFigHandle)
%     % Save the scope figure as an image
%     saveas(scopeFigHandle, 'scope_output.png');
% else
%     error('Scope figure not found.');
% end

% code to be used with multiple scopes

% model = 'PPC_scheme_04072024_a.slx';
% open_system(model);
% set_param(model, 'SimulationCommand', 'start');
% while strcmp(get_param(model, 'SimulationStatus'), 'running')
%     pause(1);
% end
% scopeBlocks = find_system(model, 'BlockType', 'Scope');
% for i = 1:length(scopeBlocks)
%     set_param([model, '/', scopeBlocks{i}], 'Open', 'on');
%     scopeHandle = findobj('Tag', 'SIMULINK_SIMSCOPE_FIGURE');
%     saveas(scopeHandle, ['scope_output_', num2str(i), '.png']);
% end
% 
% close_system(model, 0); % Close the system without saving changes


















