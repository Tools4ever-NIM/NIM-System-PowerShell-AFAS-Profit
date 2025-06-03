# version: 1.0
#
# AFAS Profit.ps1 - IDM System PowerShell Script for AFAS Profit via SOAP and REST.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#


$Log_MaskableKeys = @()


#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log verbose "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"

    if ($Connection) {
        @(
            @{
                name = 'ParticipantNr'
                type = 'textbox'
                label = 'AFAS Participant number'
                tooltip = 'Number of AFAS Profit participant'
                value = ''
            }
            @{
                name = 'ConnectionType'
                type = 'radio'
                label = 'Connection type'
                tooltip = 'Type of AFAS Profit connection'
                table = @{
                    rows = @(
                        @{ id = 'rest'; display_text = 'REST (JSON)' }
                        @{ id = 'soap'; display_text = 'SOAP (XML)' }
                    )
                    settings_radio = @{
                        value_column = 'id'
                        display_column = 'display_text'
                    }
                }
                value = 'rest'
            }
            @{
                name = 'ConnectionEnvironment'
                type = 'combo'
                label = 'Connection environment'
                tooltip = 'Environment of AFAS Profit connection'
                table = @{
                    rows = @(
                        @{ id = 'test';   display_text = 'Test' }
                        @{ id = 'accept'; display_text = 'Acceptation' }
                        @{ id = '';       display_text = 'Production' }
                    )
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'display_text'
                    }
                }
                value = 'test'
            }
            @{
                name = 'UseTls12'
                type = 'checkbox'
                label = 'Use TLS 1.2'
                value = $true
            }
            @{
                name = 'TokenVersion'
                type = 'textbox'
                label = 'Authentication token version'
                tooltip = 'Token version of AFAS Profit connection'
                value = '1'
            }
            @{
                name = 'TokenData'
                type = 'textbox'
                label = 'Authentication token data'
                tooltip = 'Token data of AFAS Profit connection'
                value = ''
            }
            @{
                name = 'Take'
                type = 'textbox'
                label = 'Result page size'
                tooltip = 'Number of rows to retrieve per request; 0 for unlimited'
                value = '0'
            }
            @{
                name = 'InventoryGetConnector'
                type = 'textbox'
                label = 'Inventory GetConnector'
                tooltip = 'GetConnector to get inventory of GetConnectors'
                value = ''
            }
        )
    }

    if ($TestConnection) {
        $connection_params = ConvertFrom-Json2 $ConnectionParams

        Log verbose "Invoke-AfasGetConnector '$($connection_params.InventoryGetConnector)'"

        Invoke-AfasGetConnector @connection_params -GetConnector $connection_params.InventoryGetConnector | Out-Null
    }

    if ($Configuration) {
        @()
    }

    Log verbose "Done"
}


#
# CRUD functions
#

$PrimaryKeys = @{}


function Idm-Dispatcher {
    param (
        # Optional Class/Operation
        [string] $Class,
        [string] $Operation,
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-Class='$Class' -Operation='$Operation' -GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($Class -eq '') {

        if ($GetMeta) {
            #
            # All Get-connectors support 'Read' operations
            #

            $system_params = ConvertFrom-Json2 $SystemParams

            Log verbose "Invoke-AfasGetConnector '$($system_params.InventoryGetConnector)'"

            Invoke-AfasGetConnector @system_params -GetConnector $system_params.InventoryGetConnector |
                ForEach-Object {
                    [ordered]@{
                        Class = $_.Naam
                        Operation = 'Read'
                        Description = $_.Omschrijving
                        Blocked = $_.Geblokkeerd
                    }
                }

            #
            # Update-connectors support 'Create' / 'Update' / 'Delete' operations
            #

            (Invoke-AfasUpdateConnector @system_params -EndPoint 'metainfo').updateConnectors |
                ForEach-Object {
                    if ($_.id -eq 'KnEmployee') {
                        [ordered]@{
                            Class = $_.id
                            Operation = 'Update'
                        }
                    }
                    else {
                        [ordered]@{
                            Class = $_.id
                            Operation = 'Create'
                        }

                        [ordered]@{
                            Class = $_.id
                            Operation = 'Update'
                        }

                        [ordered]@{
                            Class = $_.id
                            Operation = 'Delete'
                        }
                    }
                }

        }
        else {
            # Purposely no-operation.
        }

    }
    else {

        if ($GetMeta) {
            #
            # Get meta data
            #

            $system_params = ConvertFrom-Json2 $SystemParams

            Log verbose "Invoke-AfasUpdateConnector 'metainfo/update/$($Class)'"

            $response = Invoke-AfasUpdateConnector @system_params -EndPoint "metainfo/update/$($Class)"

            if ($Class -eq 'KnEmployee') {

                switch ($Operation) {

                    'Update' {
                        @{
                            semantics = 'update'
                            parameters = @(
                                ($response.fields | Where-Object { $_.primaryKey }) | ForEach-Object {
                                    @{
                                        name = $_.fieldId
                                        description = $_.label
                                        allowance = 'mandatory'
                                    }
                                }

                                ($response.objects | Where-Object { $_.name -eq 'KnPerson' }).fields | ForEach-Object {
                                    @{
                                        name = "KnPerson:$($_.fieldId)"
                                        description = $_.label
                                        allowance =
                                            switch ($_.fieldId) {
                                                'MatchPer' { 'prohibited'; break }
                                                'BcCo'     { 'mandatory';  break }
                                                default    { 'optional';   break }
                                            }
                                    }
                                }
                            )
                        }
                        break
                    }
                }

            }
            else {

                switch ($Operation) {
                    'Create' {
                        @{
                            semantics = 'create'
                            parameters = @(
                                $response.fields | ForEach-Object {
                                    @{
                                        name = $_.fieldId
                                        description = $_.label
                                        allowance = if ($_.primaryKey -or $_.mandatory) { 'mandatory' } else { 'optional' }
                                    }
                                }
                            )
                        }
                        break
                    }

                    'Read' {
                        @(
                            # No parameter items
                        )
                        break
                    }

                    'Update' {
                        @{
                            semantics = 'update'
                            parameters = @(
                                $response.fields | ForEach-Object {
                                    @{
                                        name = $_.fieldId
                                        description = $_.label
                                        allowance = if ($_.primaryKey -or $_.mandatory) { 'mandatory' } else { 'optional' }
                                    }
                                }
                            )
                        }
                        break
                    }

                    'Delete' {
                        @{
                            semantics = 'delete'
                            parameters = @(
                                $response.fields | ForEach-Object {
                                    @{
                                        name = $_.fieldId
                                        description = $_.label
                                        allowance = if ($_.primaryKey -or $_.mandatory) { 'mandatory' } else { 'optional' }
                                    }
                                }
                            )
                        }
                        break
                    }
                }

            }

        }
        else {
            #
            # Execute function
            #

            $system_params = ConvertFrom-Json2 $SystemParams

            if ($Class -eq 'KnEmployee') {

                if ($Operation -ne 'Read') {
                    if (! $Global:PrimaryKeys[$Class]) {
                        $fields = (Invoke-AfasUpdateConnector @system_params -EndPoint "metainfo/update/$($Class)").fields

                        $Global:PrimaryKeys[$Class] = @(($fields | Where-Object { $_.primaryKey }).fieldId)
                    }

                    $function_params = ConvertFrom-Json2 $FunctionParams

                    $body = @{
                        AfasEmployee = @{
                            Element = @{
                                # @EmId = "JacquelineN"
                                Fields = @{
                                }
                                Objects = @(
                                )
                            }
                        }
                    }

                    $objects = @{}

                    $function_params.Keys | ForEach-Object {
                        if ($_ -in $Global:PrimaryKeys[$Class]) {
                            $body.AfasEmployee.Element."@$_" = $function_params[$_]
                        }
                        else {
                            $field_name_components = $_.Split(':')

                            if ($field_name_components.Count -gt 2) {
                                throw "Maximum depth of sub-classes exceeded by field '$_'"
                            }
                            
                            if ($field_name_components.Count -eq 1) {
                                $body.AfasEmployee.Element.Fields."$_" = $function_params[$_]
                            }
                            else {
                                if (! $objects.($field_name_components[0])) {
                                    $objects.($field_name_components[0]) = @{
                                        Element = @{
                                            Fields = @{
                                                MatchPer = '0'
                                            }
                                        }
                                    }
                                }

                                $objects.($field_name_components[0]).Element.Fields.($field_name_components[1]) = $function_params[$_]
                            }
                        }
                    }

                    $objects.Keys | ForEach-Object {
                        $body.AfasEmployee.Element.Objects += @{
                            $_ = $objects[$_]
                        }
                    }
                }

                switch ($Operation) {
                    'Update' {
                        LogIO info "Invoke-AfasUpdateConnector" -In @system_params -Method 'PUT' -EndPoint "connectors/$($Class)" -Body $body
                            $rv = Invoke-AfasUpdateConnector @system_params -Method 'PUT' -EndPoint "connectors/$($Class)" -Body $body
                        LogIO info "Invoke-AfasUpdateConnector" -Out $rv

                        $rv
                        break
                    }
                }

            }
            else {

                if ($Operation -ne 'Read') {
                    if (! $Global:PrimaryKeys[$Class]) {
                        $fields = (Invoke-AfasUpdateConnector @system_params -EndPoint "metainfo/update/$($Class)").fields

                        $Global:PrimaryKeys[$Class] = @(($fields | Where-Object { $_.primaryKey }).fieldId)
                    }

                    $function_params = ConvertFrom-Json2 $FunctionParams

                    $body = @{
                        $Class = @{
                            Element = @{
                                Fields = @{
                                }
                            }
                        }
                    }

                    $function_params.Keys | ForEach-Object {
                        if ($_ -in $Global:PrimaryKeys[$Class]) {
                            $body.$Class.Element."@$_" = $function_params[$_]
                        }
                        else {
                            $body.$Class.Element.Fields."$_" = $function_params[$_]
                        }
                    }
                }

                switch ($Operation) {
                    'Create' {
                        LogIO info "Invoke-AfasUpdateConnector" -In @system_params -Method 'POST' -EndPoint "connectors/$($Class)" -Body $body
                            $rv = Invoke-AfasUpdateConnector @system_params -Method 'POST' -EndPoint "connectors/$($Class)" -Body $body
                        LogIO info "Invoke-AfasUpdateConnector" -Out $rv

                        $rv
                        break
                    }

                    'Read' {
                        LogIO info "Invoke-AfasGetConnector" -In @system_params -GetConnector $Class
                        Invoke-AfasGetConnector @system_params -GetConnector $Class
                        break
                    }

                    'Update' {
                        LogIO info "Invoke-AfasUpdateConnector" -In @system_params -Method 'PUT' -EndPoint "connectors/$($Class)" -Body $body
                            $rv = Invoke-AfasUpdateConnector @system_params -Method 'PUT' -EndPoint "connectors/$($Class)" -Body $body
                        LogIO info "Invoke-AfasUpdateConnector" -Out $rv

                        $rv
                        break
                    }

                    'Delete' {
                        LogIO info "Invoke-AfasUpdateConnector" -In @system_params -Method 'DELETE' -EndPoint "connectors/$($Class)" -Body $body
                            $rv = Invoke-AfasUpdateConnector @system_params -Method 'DELETE' -EndPoint "connectors/$($Class)" -Body $body
                        LogIO info "Invoke-AfasUpdateConnector" -Out $rv

                        $rv
                        break
                    }
                }

            }

        }

    }

    Log verbose "Done"
}


#
# Helper functions
#

function Invoke-AfasGetConnector {
    param(
        [parameter(ValueFromRemainingArguments=$true)] $SystemParams
    )

    function Invoke-AfasGetConnector-REST {
        param (
            [string] $ParticipantNr,
            [string] $ConnectionType,
            [string] $ConnectionEnvironment,
            [switch] $UseTls12,
            [string] $TokenVersion,
            [string] $TokenData,
            [int]    $Take,
            [string] $GetConnector
        )

        if ($UseTls12) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }

        $encoded_token = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("<token><version>$($TokenVersion)</version><data>$($TokenData)</data></token>"))
    
        $headers = @{
            Authorization = "AfasToken $encoded_token"
        }

        if ($Take -le 0) { $Take = -1 }
        $skip = if ($Take -eq -1) { -1 } else { 0 }

        $base_url = "https://$($ParticipantNr).$($ConnectionType)$($ConnectionEnvironment).afas.online/profitrestservices/connectors/$($GetConnector)"

        while ($true) {
            $url = $base_url + "?Skip=$($skip)&Take=$($Take)"

            Log debug "GET $url"

            $response = Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing
            if ($response.StatusCode -ne 200) { break }

            $rows = (ConvertFrom-Json $response.Content).rows
            if ($rows.Count -eq 0) { break }

            # Output data
            $rows

            if ($Take -eq -1) { break }
            $skip += $Take
        }
    }

    function Invoke-AfasGetConnector-SOAP {
        param (
            [string] $ParticipantNr,
            [string] $ConnectionType,
            [string] $ConnectionEnvironment,
            [switch] $UseTls12,
            [string] $TokenVersion,
            [string] $TokenData,
            [int]    $Take,
            [string] $GetConnector
        )

        function ConvertFrom-XmlSimpleElement {
            param (
                [Parameter(Mandatory, ValueFromPipeLine)] [System.Xml.XmlElement] $XmlElement,
                [object[]] $Properties
            )

            process {
                $object = [ordered]@{}

                if (-not $Properties) {
                    $Properties = $XmlElement | Get-Member -MemberType Property
                }

                foreach ($p in $Properties) {
                    $object[$p.name] = $XmlElement.($p.name)
                }

                New-Object -TypeName PSObject -Property $object
            }
        }

        if ($UseTls12) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }

        if ($Take -le 0) { $Take = -1 }
        $skip = if ($Take -eq -1) { -1 } else { 0 }

        $url = "https://$($ParticipantNr).$($ConnectionType)$($ConnectionEnvironment).afas.online/profitservices/AppConnectorGet.asmx?WSDL"
        Log debug "url = [$url]"

        $webservicex = New-WebServiceProxy -Uri $url -namespace WebServiceProxy -Class AfasConnector
        $properties = $null

        while ($true) {
            [XML]$response = $webservicex.GetData("<token><version>$($TokenVersion)</version><data>$($TokenData)</data></token>", $GetConnector, '', $skip, $Take)

            if (-not $properties) {
                $properties = $response.AfasGetConnector.schema.element.complexType.choice.element.complexType.sequence.element | ConvertFrom-XmlSimpleElement
            }

            $xml_data = $response.AfasGetConnector.$GetConnector
            if ($xml_data.Count -eq 0) { break }

            # Output data
            $xml_data | ConvertFrom-XmlSimpleElement -Properties $properties

            if ($Take -eq -1) { break }
            $skip += $Take
        }
    }

    $system_params = @{}

    for ($i = 0; $i -lt $SystemParams.Count; $i += 2) {
        $system_params[($SystemParams[$i]-replace '^-'-replace ':$')] = $SystemParams[$i + 1]
    }

    try {
        Invoke-Expression "Invoke-AfasGetConnector-$($system_params.ConnectionType) @system_params"
    }
    catch {
        Log error "Failed: $_ ($($_.Exception.Response.ResponseUri))"
        Write-Error $_
    }
}


function Invoke-AfasUpdateConnector {
    param(
        [parameter(ValueFromRemainingArguments=$true)] $SystemParams
    )

    function Invoke-AfasUpdateConnector-REST {
        param (
            [string] $ParticipantNr,
            [string] $ConnectionType,
            [string] $ConnectionEnvironment,
            [switch] $UseTls12,
            [string] $TokenVersion,
            [string] $TokenData,
            [string] $Method,
            [string] $EndPoint,
            [object] $Body
        )

        if ($UseTls12) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }

        $encoded_token = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("<token><version>$($TokenVersion)</version><data>$($TokenData)</data></token>"))

        $args = @{
            Headers = @{
                Authorization = "AfasToken $encoded_token"
            }
            Method = if ($Method) { $Method } else { 'GET' }
            Uri = "https://$($ParticipantNr).$($ConnectionType)$($ConnectionEnvironment).afas.online/profitrestservices/$($EndPoint)"
            UseBasicParsing = $true
        }

        if ($Body) {
            $args.Body = (ConvertTo-Json $Body -Depth 32)
        }

        Log debug "$($args.Method) $($args.Uri)"

        $response = Invoke-WebRequest @args
        if ($response.StatusCode -ne 200) { return }

        # Output data
        ConvertFrom-Json $response.Content
    }

    function Invoke-AfasUpdateConnector-SOAP {
        param (
            [string] $ParticipantNr,
            [string] $ConnectionType,
            [string] $ConnectionEnvironment,
            [switch] $UseTls12,
            [string] $TokenVersion,
            [string] $TokenData,
            [int]    $Take,
            [string] $Method,
            [string] $EndPoint,
            [object] $Body
        )

        function ConvertFrom-XmlSimpleElement {
            param (
                [Parameter(Mandatory, ValueFromPipeLine)] [System.Xml.XmlElement] $XmlElement,
                [object[]] $Properties
            )

            process {
                $object = [ordered]@{}

                if (-not $Properties) {
                    $Properties = $XmlElement | Get-Member -MemberType Property
                }

                foreach ($p in $Properties) {
                    $object[$p.name] = $XmlElement.($p.name)
                }

                New-Object -TypeName PSObject -Property $object
            }
        }

        if ($UseTls12) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }

        if ($Take -le 0) { $Take = -1 }
        $skip = if ($Take -eq -1) { -1 } else { 0 }

        $url = "https://$($ParticipantNr).$($ConnectionType)$($ConnectionEnvironment).afas.online/profitservices/AppConnectorGet.asmx?WSDL"
        Log debug "url = [$url]"

        $webservicex = New-WebServiceProxy -Uri $url -namespace WebServiceProxy -Class AfasConnector
        $properties = $null

        while ($true) {
            [XML]$response = $webservicex.GetData("<token><version>$($TokenVersion)</version><data>$($TokenData)</data></token>", $GetConnector, '', $skip, $Take)

            if (-not $properties) {
                $properties = $response.AfasGetConnector.schema.element.complexType.choice.element.complexType.sequence.element | ConvertFrom-XmlSimpleElement
            }

            $xml_data = $response.AfasGetConnector.$GetConnector
            if ($xml_data.Count -eq 0) { break }

            # Output data
            $xml_data | ConvertFrom-XmlSimpleElement -Properties $properties

            if ($Take -eq -1) { break }
            $skip += $Take
        }
    }

    $system_params = @{}

    for ($i = 0; $i -lt $SystemParams.Count; $i += 2) {
        $system_params[($SystemParams[$i]-replace '^-'-replace ':$')] = $SystemParams[$i + 1]
    }

    try {
        Invoke-Expression "Invoke-AfasUpdateConnector-$($system_params.ConnectionType) @system_params"
    }
    catch {
        Log error "Failed: $_ ($($_.Exception.Response.ResponseUri))"
        Write-Error $_
    }
}
