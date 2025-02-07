%DESARROLLADOR: JUAN ORTIZ (2024)
%EMPRESA: CONSORCIO SIG-ELECTRIC
%DESCRIPCION: SCRIPT PARA COMPARAR EXCEL BASE Y LEVANTAMIENTO EN CAMPO
%VERSION 1

clc
clear 
format LONGG

addpath('F_Funciones\');
inputFileBase='H:\Mi unidad\ProyectoSurveyMatlab\ArchivosBase\';
addpath(inputFileBase);
folderOut='H:\Mi unidad\ProyectoSurveyMatlab\Analisis_SurveyMatlab';
addpath(folderOut);


%%                   CAMBIOS IMPORTANTES 
%Estos cambios se deben hacer en caso que se reemplace el archivo de
%ingreso o salida de datos, cambios de variable y hojas en excel

%%%%%%%%%%%%%%%%%%%%%%%Ingreso de datos%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Alimentador='59C';  
ExcelInputBase = [inputFileBase,'VALIDACION_59C.xlsx'];
folderOutxAlim = [folderOut,'\Resultados' ,Alimentador,'\'];% Directorio para la salida
ExcelInputAnalisis = [folderOutxAlim,'SuperTabla59C_30-Jan-2025.xlsx']; 
ExcelOutputAlertas = [folderOutxAlim,'Alerts59C_30-Jan-2025.xlsx']; 

    % BASE Hojas de lectura
    % sheetInBase1='TRAFO';
    % sheetInBase2='POSTE';
    sheetInBase4='PUNTO DE CARGA';
    % sheetInBase5='MEDIDORES';
                                                            
    % CAMPO Hojas de lectura
    sheetInCampo2='PuntoCarga_Poste_Medidor_CAMPO';%Cambia segun el nombre del formulario 

%%%%%%%%%%%%%%%%%%%%%% Salida de datos %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Obtener información del archivo
%Hojas de impresion
% sheetOut1='Trafos_CAMPO';
sheetOut2='PuntoCarga_Poste_Medidor_CAMPO';
sheetOut3='Alertas';


rangePrintFirst=1; 

if ~exist(folderOutxAlim, 'dir')
    mkdir(folderOutxAlim); % Crear la carpeta si no existe
end

%%                   LECTURA DE DATOS 

%%%%%%%%%%%%%%%%%%%%%%%%% BASE %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% 

%++++++++++++++++PCarga++++++++++++++++++++++++++++++++++++++++++++++++++++
PCargaBase = F_leerDatosExcel(ExcelInputBase, sheetInBase4); %Hoja completa de transformadores en la Base del Alimentador

%%%%%%%%%%%%%%%%%%%%%%%%% Analisis %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%++++++++++++++++ Medidores +++++++++++++++++++++++++++++++++++++++++++++++
Medidores_Campo= F_leerDatosExcel(ExcelInputAnalisis, sheetInCampo2);

%%                        ALERTAS

% Paso 1: Extraer latitud y longitud
Lat = Medidores_Campo.Lat_PCrgaCampo; % Columna de latitudes
Lon = Medidores_Campo.Long_PCargaCampo; % Columna de longitudes


% Verificar si las columnas son celdas, y convertirlas a valores numéricos
if iscell(Lat)
    Lat = cellfun(@str2double, Lat); % Convertir cada elemento de la celda a double
end
if iscell(Lon)
    Lon = cellfun(@str2double, Lon); % Convertir cada elemento de la celda a double
end

% Paso 2: Convertir coordenadas geográficas a UTM usando deg2utm
[x, y, utmzone] = deg2utm_ecuador(Lat, Lon);

% Paso 3: Agregar las coordenadas X e Y como nuevas columnas en la tabla
Medidores_Campo.X_UTM_PcBase = x; % Coordenada UTM X
Medidores_Campo.Y_UTM_PcBase = y; % Coordenada UTM Y
Medidores_Campo.UTM_Zone = utmzone; % Zona UTM (opcional)

% Formatear las columnas numéricas
Medidores_Campo.X_UTM_PcBase = arrayfun(@(val) sprintf('%.6f', val), Medidores_Campo.X_UTM_PcBase, 'UniformOutput', false);
Medidores_Campo.Y_UTM_PcBase = arrayfun(@(val) sprintf('%.6f', val), Medidores_Campo.Y_UTM_PcBase, 'UniformOutput', false);

% Crear una nueva columna en PC_Campo para almacenar las distancias
Medidores_Campo.Distancia = nan(height(Medidores_Campo), 1); % Preasignar con NaN
Medidores_Campo.AlertaDesplazamiento = repmat({''}, height(Medidores_Campo), 1);

% Recorrer cada fila de PC_Campo
for i = 1:height(Medidores_Campo)
    % Extraer el valor de PC_Campo.PCargaAsignado (asegúrate de que no sea celda)
    valorAsignado = Medidores_Campo.PCargaAsignado{i}; % Si es celda, extraer el contenido
    
    % Buscar coincidencias entre PC_Campo.PCargaAsignado y PCargaBase.OID_PC
    indiceBase = find(strcmp(PCargaBase.OID_PC, valorAsignado)); % Usar strcmp para comparar celdas
    
    if ~isempty(indiceBase) % Si hay coincidencias
        % Extraer coordenadas de la base
        coordX_Base = str2double(PCargaBase.COORD_X(indiceBase));
        coordY_Base = str2double(PCargaBase.COORD_Y(indiceBase));
        
        % Extraer coordenadas de campo
        coordX_Campo = cellfun(@str2double, Medidores_Campo.X_UTM_PcBase(i));
        coordY_Campo = cellfun(@str2double, Medidores_Campo.Y_UTM_PcBase(i));

        % Asegurarse de que las coordenadas sean numéricas
        if ~any(isnan(coordX_Campo)) || ~any(isnan(coordY_Campo))
        % Calcular la distancia euclidiana
        distancia = sqrt((coordX_Campo - coordX_Base)^2 + (coordY_Campo - coordY_Base)^2);
        
        % Almacenar la distancia en la tabla PC_Campo
        Medidores_Campo.Distancia(i) = distancia;

                % Verificar si la distancia es mayor a 10
        if distancia > 10
            Medidores_Campo.AlertaDesplazamiento(i) = {'A PC Desplazado'}; % Agregar alerta
        end 
        if distancia > 100
            Medidores_Campo.AlertaDesplazamiento(i) = {'B Revisar Ubicacion'}; % Agregar alerta
        end

       end
   end
end

Medidores_Campo = Medidores_Campo(:, {'PCargaCampo', 'PCargaBase', 'PCargaAsignado','PcAsignado_GlobalID','ConductorAcometida' ,'SugerenciaMatlabPoste','PosteCampoPcA','PosteBasePcA','MDENUMEMP','MedidoresQRyTxt','SugerenciaMatlabMedidores','AlertaDesplazamiento','CodigoPuesto','ComentariosCampo','Comentarios_PCarga','Comentarios_Poste','Comentarios_Trafo','X_UTM_PcBase','Y_UTM_PcBase','Distancia','Long_PosteCampo','Lat_PosteCampo','Long_TrafoCampo','Lat_TrafoCampo'}); % Reorganizar con la nueva columna


%%                  IMPRIMIR RESULTADOS

writetable(Medidores_Campo,ExcelInputAnalisis,'Sheet', sheetOut2,'Range','A1');%Imprime tabla de Alertas

rowsAlertas = ~cellfun(@isempty, Medidores_Campo.SugerenciaMatlabPoste) | ~cellfun(@isempty, Medidores_Campo.SugerenciaMatlabMedidores) ;
AlertasMedidores= Medidores_Campo(rowsAlertas, :);

writetable(AlertasMedidores,ExcelOutputAlertas,'Sheet', sheetOut3,'Range','A1');%Imprime tabla de Alertas

%%                          FUNCIONES

function [x, y, utmzone] = deg2utm_ecuador(Lat, Lon)
    % Calcula coordenadas UTM para coordenadas geográficas en Ecuador.
    %
    % Entrada:
    %   Lat - Vector de latitudes (en grados).
    %   Lon - Vector de longitudes (en grados).
    %
    % Salida:
    %   x, y    - Coordenadas UTM (en metros).
    %   utmzone - Código UTM (zona + hemisferio).

    % Asegurarnos de que las dimensiones coincidan
    if length(Lat) ~= length(Lon)
        error('Lat y Lon deben tener el mismo tamaño.');
    end

    % Inicializar los vectores de salida
    x = zeros(size(Lat));
    y = zeros(size(Lat));
    utmzone = strings(size(Lat));  % Cadena para almacenar la zona UTM

    % Recorrer las coordenadas y calcular UTM
    for i = 1:length(Lat)
        % Determinar el hemisferio y EPSG
        if Lat(i) >= 0
            % Hemisferio Norte
            proj = projcrs(32717); % EPSG 32617 (UTM Zona 17N, WGS84)
            hemisphere = 'N';
        else
            % Hemisferio Sur
            proj = projcrs(32717); % EPSG 32717 (UTM Zona 17S, WGS84)
            hemisphere = 'S';
        end

        % Calcular las coordenadas UTM con projfwd
        [x(i), y(i)] = projfwd(proj, Lat(i), Lon(i));

        % Crear el código de zona UTM
        utmzone(i) = sprintf('17 %s', hemisphere);
    end
end
%%                             NOTAS 
% %%%%%%%%%%%%%%%%%%%%%%%%% TRAFOS %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


