function progetto_software_main_240911(modelname)
% Define the folder where images will be saved
fullFilePath=mfilename('fullpath');
currentFileDirectory=fileparts(fullFilePath);
% attach a subdirectory to the current file directory
newDirectory=fullfile(currentFileDirectory,'Pictures');
saveFolder=newDirectory;
% Create the folder if it doesn't exist
if ~exist(saveFolder,'dir')
    mkdir(saveFolder);
end
modelnameStr=convertCharsToStrings(modelname)
open_system(modelnameStr);
sim(modelnameStr);
model = modelnameStr;
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
    ImagePath=fullfile(saveFolder,fileName);
    %saveas(figureHandle,fileName);
    saveas(figureHandle,ImagePath)
end
end


    


















