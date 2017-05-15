% v1.3

clear
clc

cd 'C:\Users\Daniel\Dropbox\Greenvis\_Projecten\Warmtekansen spotten\11_Data_analyse\14_Utrecht_LWD_hybride\n01_Energie_in_beeld_data\PerProvincie'
dirInfo = dir;
dirInfo(1:2) = [];
nProv = length(dirInfo);

gemeenteList = {};
PClist = {};
gasverbruikList_m3 = {};
aantalAanslList = {};
for iProv = 1:length(dirInfo)
    cd(dirInfo(iProv).name)
    provDirInfo = dir;
    provDirInfo(1:2) = [];
    nGemeentesInProv = length(provDirInfo);
    
    for iGemeente = 1:length(provDirInfo)
        fileName = provDirInfo(iGemeente).name;
        % Find index of 2016 sheet
        [status,sheets] = xlsfinfo(fileName);
        index_2016 = find(contains(sheets,'2016'));
        
        [~,~,PClistNew] = xlsread(fileName,index_2016,'B2:B1000');
        PClistNew = PClistNew(cellfun(@ischar,PClistNew));
        PClist = [PClist;PClistNew];
        nPCinGem = length(PClistNew);
        
        [~,~,gemeenteListNew] = xlsread(fileName,index_2016,['A2:A' num2str(1+nPCinGem)]);
        gemeenteList = [gemeenteList;gemeenteListNew];
        
        [~,~,aantalAanslListNew] = xlsread(fileName,index_2016,['E2:E' num2str(1+nPCinGem)]);
        aantalAanslList = [aantalAanslList;aantalAanslListNew];
        
        [~,~,gasverbruikListNew_m3] = xlsread(fileName,index_2016,['H2:H' num2str(1+nPCinGem)]);
        gasverbruikList_m3 = [gasverbruikList_m3;gasverbruikListNew_m3];
        
        disp(['nPC: ' num2str(length(PClistNew)) ' - nGem: ' num2str(length(gemeenteListNew)) ...
            ' - nAantalAansl: ' num2str(length(aantalAanslListNew)) ' - nGasv: ' num2str(length(gasverbruikListNew_m3))])
        disp(['Gemeente ' num2str(iGemeente) '/' num2str(nGemeentesInProv)])
    end
    cd ..
    disp(['Provincie ' num2str(iProv) '/' num2str(nProv) '\n'])
end

nPCtot = length(PClist);

outputPath = 'C:\Users\Daniel\Dropbox\Greenvis\_Projecten\Warmtekansen spotten\11_Data_analyse\14_Utrecht_LWD_hybride\n01_Energie_in_beeld_data\Nederland\Matlab\samenvatting.xlsx';
xlswrite(outputPath,gemeenteList,['A2:A' num2str(nPCtot+1)])
xlswrite(outputPath,PClist,['B2:B' num2str(nPCtot+1)])
xlswrite(outputPath,aantalAanslList,['C2:C' num2str(nPCtot+1)])
xlswrite(outputPath,gasverbruikList_m3,['D2:D' num2str(nPCtot+1)])
