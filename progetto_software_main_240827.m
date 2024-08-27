% multiple scopes

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
    % Retrieve the handle of the subsystem
    scopesubsys = get_param(scopeBlocks{i}, 'Parent');
    % get all figure handles
    figureHandle = findall(0, 'Type', 'figure', 'Name', get_param(scopeHandle, 'Name'));
    
    % Save the scope figure as an image    
    scopename = get_param(scopeHandle, 'Name');
    slashposition=strfind(scopesubsys,'/');
    length_string=strlength(scopesubsys);
    parts = strsplit(scopesubsys, '/');
    if length(parts)>2
        firstpart=parts(length(parts)-1);
        secondpart=parts(length(parts));
        firstpartchar=char(firstpart);
        secondpartchar=char(secondpart);
        fileName = strcat(scopename,'_',firstpartchar,'_',secondpartchar,'.png');
    else
        firstpart=''
        secondpart=parts(length(parts));
        secondpartchar=char(secondpart);
        fileName = strcat(scopename,'_',secondpartchar,'.png');
    end
    %extractedPortion = scopesubsys((slashposition+1):length_string);
    %fileName=strcat(scopename,'.png')
    %fileName = strcat(scopename,'_',extractedPortion,'.png');
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


















