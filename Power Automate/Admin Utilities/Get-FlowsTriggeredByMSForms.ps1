Add-PowerAppsAccount

$formId = "fgosTqeRf0ybo164HIjLj1IeUikwJpVFoaicVhb1rGJUODUyVkYyTkkzUlc0SldISDREUUw3TlJQTiQlQCN0PWcu"

$environments = Get-AdminPowerAppEnvironment

foreach ($environment in $environments) {

    $allFlows  = Get-AdminFlow -EnvironmentName $environment.EnvironmentName

    foreach ($flow in $allFlows) {
        Write-Host $flow.DisplayName
        foreach ($trigger in $flow.Internal.properties.definitionSummary.triggers) {
            if ($trigger.swaggerOperationId -eq "CreateFormWebhook") {
                Write-Host "Form Trigger Found $($flow.DisplayName)" -ForegroundColor DarkYellow
                $flowDef = Get-Flow -EnvironmentName $environment.EnvironmentName -FlowName $flow.FlowName
                if ($flowDef.Internal.properties.definition.triggers.When_a_new_response_is_submitted.inputs.parameters.form_id -eq $formId) {
                    Write-Host "Found flow that connects to form $($formId)" -ForegroundColor Green
                    Write-Host "Environment: $($environment.DisplayName)" -ForegroundColor Green
                    Write-Host "Flow Name: $($flow.DisplayName)" -ForegroundColor Green
                    Write-Host "Owner: $($flow.CreatedBy.userId)" -ForegroundColor Green
                }
            }
        }
    }
}