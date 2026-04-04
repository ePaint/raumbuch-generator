@{
    TemplateFile = 'Input/Raumbuch_Vorlage_alt.docx'
    OutputFolder = 'Output'

    # Data source: 'Excel' or 'API'
    DataSource   = 'Excel'

    # Excel settings (when DataSource = 'Excel')
    Excel = @{
        DataFile       = 'Input/ZB3.0.xlsx'
        RoomCodeColumn = 'Code'
    }

    # API settings (when DataSource = 'API')
    API = @{
        EndpointFile  = 'temp/addon/API Call.txt'
        KeyFile       = 'api-key.txt'
        RoomCodeField = 'room_func_no'
    }

    # Value replacements (case-insensitive, use lowercase keys)
    ValueMap = @{
        'true'  = 'ja'
        'false' = 'nein'
    }
}
