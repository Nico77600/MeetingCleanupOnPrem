#Requires -Version 7.4
# Signature-dependent cases carry the RealSignature tag: they need the owner-signed release package named by
# MCO_RELEASE_PACKAGE and are skipped without it. The other cases run the real verifier on a fixture copy whose
# Authenticode check is replaced by a no-op, so that the manifest, the inventory, the hashes and the folder rules are
# exercised before the owner signs.
BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Source = Join-Path $script:RepoRoot 'package'
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'
    $script:Signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
    $script:FixtureNumber = 0
    $script:Fixture = $null

    # The manifest writer of the package builder, used as is.
    $builderPath = Join-Path $script:RepoRoot 'tools\New-MeetingCleanupOnPremPackage.ps1'
    $ast = [Management.Automation.Language.Parser]::ParseFile($builderPath, [ref]$null, [ref]$null)
    $signer = $script:Signer
    $version = (Import-PowerShellDataFile -LiteralPath (Join-Path $script:Source 'MeetingCleanupOnPrem.psd1')).ModuleVersion
    foreach ($name in 'Get-RelativeFile', 'Write-Manifest') {
        $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name }, $true)
        if ($null -eq $function) { throw "The builder function $name is missing." }
        . ([scriptblock]::Create($function.Extent.Text))
    }

    function New-FixtureTemplate {
        $template = Join-Path $TestDrive 'template'
        Copy-Item -LiteralPath $script:Source -Destination $template -Recurse
        foreach ($folder in 'reports', 'logs') {
            Remove-Item -LiteralPath (Join-Path $template $folder) -Recurse -Force -ErrorAction SilentlyContinue
        }
        Remove-Item -LiteralPath (Join-Path $template 'PackageIntegrity.psd1') -ErrorAction SilentlyContinue
        # A release package holds the HTML guides only.
        Remove-Item -LiteralPath (Join-Path $template 'docs\images') -Recurse -Force
        Get-ChildItem -LiteralPath (Join-Path $template 'docs') -Filter '*.md' | Remove-Item
        Copy-Item -LiteralPath (Join-Path $script:RepoRoot 'CHANGELOG.md') -Destination $template
        # The test copy of the verifier does not check Authenticode: everything else is the delivered code.
        $verifier = Join-Path $template 'src\Test-PackageIntegrity.ps1'
        $text = [IO.File]::ReadAllText($verifier)
        $stubbed = [regex]::Replace($text, '(?s)function Assert-OwnerSignature\(\[string\]\$Path\) \{.*?\r?\n\}\r?\n', "function Assert-OwnerSignature([string]`$Path) { }`r`n")
        if ($stubbed -ceq $text) { throw 'The owner-signature check was not found in the verifier.' }
        [IO.File]::WriteAllText($verifier, $stubbed, [Text.UTF8Encoding]::new($false))
        [void](Write-Manifest $template)
        $template
    }
    function Reset-Fixture([switch]$Source) {
        Remove-Module MeetingCleanupOnPrem -Force -ErrorAction SilentlyContinue
        $script:FixtureNumber++
        $script:Fixture = Join-Path $TestDrive ('package-{0:D4}' -f $script:FixtureNumber)
        Copy-Item -LiteralPath $script:Template -Destination $script:Fixture -Recurse
        # A source checkout: no file copied at build time.
        if ($Source) { Remove-Item -LiteralPath (Join-Path $script:Fixture 'CHANGELOG.md') }
    }
    function Invoke-Guard([string[]]$Switch = @()) {
        $verifier = Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1'
        $arguments = @{ Root = $script:Fixture }
        foreach ($name in $Switch) { $arguments[$name] = $true }
        & $verifier @arguments 3> $null
    }
    function Invoke-Child([string]$Command) {
        $output = & $script:Pwsh -NoProfile -NonInteractive -Command $Command 2>&1 | ForEach-Object { "$_" }
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Text = ($output -join "`n") }
    }
    function Edit-Protected([string]$Path) {
        if ([IO.Path]::GetExtension($Path) -in '.ps1', '.psm1', '.psd1') {
            [IO.File]::WriteAllText($Path, "# Evidence-only tamper`r`n" + [IO.File]::ReadAllText($Path), [Text.UTF8Encoding]::new($true))
        } else { [IO.File]::AppendAllBytes($Path, [Text.Encoding]::ASCII.GetBytes("`r`n# Evidence-only tamper")) }
    }
    $script:Template = New-FixtureTemplate
    $manifest = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Template 'PackageIntegrity.psd1')
    $script:Entries = @($manifest.Files)
    $script:NoConfig = Join-Path $TestDrive 'no-such-config.psd1'
}

Describe 'Signed manifest and verifier' {
    BeforeEach { Reset-Fixture }
    It 'accepts the distribution, the same files as a source checkout, and each in its own mode' {
        (Invoke-Guard @('Distribution')).Status | Should -Be 'Valid'
        (Invoke-Guard).Status | Should -Be 'Valid (files only)'
        (Invoke-Guard @('DevelopmentMode')).Status | Should -Be 'Valid (DevelopmentMode)'
        Reset-Fixture -Source
        (Invoke-Guard).Status | Should -Be 'Valid (files only)'
        (Invoke-Guard @('DevelopmentMode')).Status | Should -Be 'Valid (DevelopmentMode)'
        { Invoke-Guard @('Distribution') } | Should -Throw
    }
    It 'refuses -Distribution together with -DevelopmentMode' {
        { Invoke-Guard @('Distribution', 'DevelopmentMode') } | Should -Throw -ExpectedMessage '*cannot be combined*'
    }
    It 'inventories exactly the protected files and not the editable configuration' {
        $paths = @($script:Entries.Path)
        $paths.Count | Should -Be 22
        $paths | Should -Not -Contain 'config\MeetingCleanupOnPrem.config.psd1'
        $paths | Should -Not -Contain 'PackageIntegrity.psd1'
        foreach ($relative in 'Invoke-MeetingCleanupOnPrem.ps1', 'MeetingCleanupOnPrem.psm1', 'MeetingCleanupOnPrem.psd1', 'src\Test-PackageIntegrity.ps1',
            'src\MeetingCleanupOnPrem.Native.cs', 'templates\Report.template.html', 'docs\MeetingCleanupOnPrem-UserGuide.html', 'docs\MeetingCleanupOnPrem-Guide.html',
            'README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md', 'CHANGELOG.md') { $paths | Should -Contain $relative }
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $script:Source 'src') -File -Filter '*.ps1') { $paths | Should -Contain "src\$($file.Name)" }
        @($paths | Sort-Object -CaseSensitive) | Should -Be $paths
    }
    It 'keeps the verifier only in src and authenticates the package folder rather than src' {
        Copy-Item -LiteralPath (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1') -Destination (Join-Path $script:Fixture 'Test-PackageIntegrity.ps1')
        { & (Join-Path $script:Fixture 'Test-PackageIntegrity.ps1') -Root $script:Fixture } | Should -Throw -ExpectedMessage '*verifier must belong to this package root*'
        { & (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1') -Root (Join-Path $script:Fixture 'src') } | Should -Throw
    }
    It 'accepts the editable configuration and the generated folders' {
        Add-Content (Join-Path $script:Fixture 'config\MeetingCleanupOnPrem.config.psd1') '# Fictitious customization'
        foreach ($folder in 'reports', 'logs') {
            [void][IO.Directory]::CreateDirectory((Join-Path $script:Fixture $folder))
            Set-Content (Join-Path $script:Fixture "$folder\generated.txt") 'generated'
        }
        (Invoke-Guard @('Distribution')).Status | Should -Be 'Valid'
    }
    It 'rejects mutations of every protected entry' {
        foreach ($entry in $script:Entries) {
            Reset-Fixture
            Edit-Protected (Join-Path $script:Fixture $entry.Path)
            { Invoke-Guard @('Distribution') } | Should -Throw -ExpectedMessage "*SHA256 mismatch: $($entry.Path)*"
            if ($entry.Path -cne 'CHANGELOG.md') {
                { Invoke-Guard } | Should -Throw -ExpectedMessage "*SHA256 mismatch: $($entry.Path)*"
                { Invoke-Guard @('DevelopmentMode') } | Should -Throw -ExpectedMessage "*SHA256 mismatch: $($entry.Path)*"
            }
        }
    }
    It 'rejects missing mandatory files' {
        foreach ($entry in $script:Entries) {
            Reset-Fixture
            Remove-Item -LiteralPath (Join-Path $script:Fixture $entry.Path)
            { Invoke-Guard @('Distribution') } | Should -Throw
            if ($entry.Path -cne 'CHANGELOG.md') { { Invoke-Guard } | Should -Throw }
        }
    }
    It 'rejects any unknown stable file' {
        foreach ($relative in 'src\extra.cs', 'src\MeetingCleanupOnPrem.Extra.ps1', 'src\extra.ps1', 'templates\extra.html', 'docs\extra.ps1', 'docs\extra.html',
            'docs\images\extra.png', 'config\extra.psd1', 'extra.ps1', 'extra.dll', 'extra.cmd', 'tools\Test-PackageIntegrity.ps1', 'Test-PackageIntegrity.ps1') {
            Reset-Fixture
            [void][IO.Directory]::CreateDirectory((Split-Path (Join-Path $script:Fixture $relative) -Parent))
            Set-Content (Join-Path $script:Fixture $relative) 'not executable'
            { Invoke-Guard } | Should -Throw -ExpectedMessage '*unexpected stable file*' -Because $relative
        }
    }
    It 'rejects a junction in a protected subtree' {
        $outside = Join-Path $TestDrive 'junction-target'
        [void][IO.Directory]::CreateDirectory($outside)
        foreach ($relative in 'src', 'templates', 'docs', 'config') {
            Reset-Fixture
            $link = Join-Path $script:Fixture $relative
            Remove-Item -LiteralPath $link -Recurse -Force
            New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
            try { { Invoke-Guard } | Should -Throw -Because $relative }
            finally { (Get-Item -LiteralPath $link -Force).Delete() }
        }
        foreach ($generated in 'reports', 'logs') {
            Reset-Fixture
            $link = Join-Path $script:Fixture $generated
            New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
            try { { Invoke-Guard } | Should -Throw -ExpectedMessage '*reparse point*' -Because $generated }
            finally { (Get-Item -LiteralPath $link -Force).Delete() }
        }
    }
    It 'rejects an absent manifest and a manifest of the wrong shape' {
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        { Invoke-Guard } | Should -Throw -ExpectedMessage '*PackageIntegrity.psd1 is missing. Use the complete release ZIP (https://github.com/Nico77600/MeetingCleanupOnPrem/releases/latest)*-DevelopmentMode*'
        Reset-Fixture
        Copy-Item -LiteralPath (Join-Path $script:Fixture 'MeetingCleanupOnPrem.psd1') -Destination (Join-Path $script:Fixture 'PackageIntegrity.psd1') -Force
        { Invoke-Guard } | Should -Throw -ExpectedMessage '*unsupported or incomplete manifest schema*'
    }
    It 'rejects a manifest that omits, duplicates, adds or retargets an entry' {
        $path = Join-Path $script:Fixture 'PackageIntegrity.psd1'
        $original = [IO.File]::ReadAllText($path)
        $variants = [ordered]@{
            omitted     = [regex]::Replace($original, "(?m)^\s*@\{ Path = 'README.md';.*\r?\n", '')
            duplicate   = $original.Replace("@{ Path = 'README.md';", "@{ Path = 'THIRD-PARTY-NOTICES.md';")
            unlisted    = $original.Replace("Path = 'README.md'", "Path = 'src\Extra.cs'")
            traversal   = $original.Replace("Path = 'README.md'", "Path = '..\README.md'")
            absolute    = $original.Replace("Path = 'README.md'", "Path = 'C:\Windows\win.ini'")
            wronghash   = $original.Replace("Path = 'README.md'; Sha256 = '", "Path = 'README.md'; Sha256 = 'g")
            wrongsigner = $original.Replace($script:Signer, '0' * 40)
        }
        foreach ($name in $variants.Keys) {
            [IO.File]::WriteAllText($path, $variants[$name], [Text.UTF8Encoding]::new($true))
            { Invoke-Guard @('Distribution') } | Should -Throw -ExpectedMessage '*Package integrity:*' -Because $name
        }
    }
}

Describe 'Integrity at the entry points' {
    BeforeEach { Reset-Fixture }
    It 'refuses to start the entry script and the module from a changed package, before anything is compiled' {
        foreach ($relative in 'src\MeetingCleanupOnPrem.Native.cs', 'src\Test-PackageIntegrity.ps1', 'MeetingCleanupOnPrem.psm1', 'Invoke-MeetingCleanupOnPrem.ps1') {
            Reset-Fixture
            Edit-Protected (Join-Path $script:Fixture $relative)
            $run = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanupOnPrem.ps1')' -ConfigPath '$($script:NoConfig)'; 'EXIT=' + `$LASTEXITCODE")
            $run.Text | Should -Match 'Package integrity' -Because $relative
            $run.Text | Should -Not -Match 'EXIT=0'
            Test-Path -LiteralPath (Join-Path $script:Fixture 'reports') | Should -BeFalse
        }
    }
    It 'refuses a tampered verifier before the module is imported' {
        Edit-Protected (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1')
        { Import-Module (Join-Path $script:Fixture 'MeetingCleanupOnPrem.psd1') -Force -ErrorAction Stop } | Should -Throw -ExpectedMessage '*Package integrity:*'
    }
    It 'does not run a verifier that lacks the owner signature, even from an otherwise correct package' {
        # The fixture verifier is unsigned: the module and the entry script must refuse it instead of executing it.
        { Import-Module (Join-Path $script:Fixture 'MeetingCleanupOnPrem.psd1') -Force -ErrorAction Stop } | Should -Throw -ExpectedMessage '*owner signature*'
        $run = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanupOnPrem.ps1')' -ConfigPath '$($script:NoConfig)' -DevelopmentMode; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'owner signature'
    }
    It 'accepts an unverified source checkout only with -DevelopmentMode' {
        Reset-Fixture -Source
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        $plain = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanupOnPrem.ps1')' -ConfigPath '$($script:NoConfig)'; 'EXIT=' + `$LASTEXITCODE")
        $plain.Text | Should -Match 'PackageIntegrity.psd1 is missing\. Use the complete release ZIP .*releases/latest.*-DevelopmentMode'
        $development = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanupOnPrem.ps1')' -ConfigPath '$($script:NoConfig)' -DevelopmentMode 3>&1 | Out-String")
        $development.Text | Should -Match 'DevelopmentMode: no package manifest'
        # The run goes on to the configuration: it is the missing configuration file that stops it.
        $development.Text | Should -Match 'no-such-config'
        $development.Text | Should -Not -Match 'positional|Package integrity: verifier|unexpected stable file'
        $development.Text | Should -Not -Match 'PackageIntegrity.psd1 is missing'
        # One warning per process: the module does not repeat it at import.
        ([regex]::Matches($development.Text, 'no package manifest|no PackageIntegrity\.psd1')).Count | Should -Be 1
    }
}

Describe 'Guides' {
    It 'regenerates the HTML guides from their Markdown sources without a difference' {
        $builder = Join-Path $script:RepoRoot 'tools\Build-Documentation.ps1'
        foreach ($name in 'MeetingCleanupOnPrem-UserGuide', 'MeetingCleanupOnPrem-Guide') {
            $target = Join-Path $TestDrive "$name.html"
            & $builder -Source (Join-Path $script:Source "docs\$name.md") -Destination $target | Out-Null
            (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $script:Source "docs\$name.html") -Algorithm SHA256).Hash -Because $name
        }
    }
    It 'pins the line endings of every hashed file' {
        $attributes = Get-Content -LiteralPath (Join-Path $script:RepoRoot '.gitattributes') -Raw
        foreach ($pattern in 'package/Invoke-MeetingCleanupOnPrem.ps1', 'package/MeetingCleanupOnPrem.psm1', 'package/MeetingCleanupOnPrem.psd1', 'package/src/*.ps1', 'package/PackageIntegrity.psd1',
            'package/src/*.cs', 'package/templates/Report.template.html', 'package/docs/MeetingCleanupOnPrem-Guide.html', 'package/docs/MeetingCleanupOnPrem-UserGuide.html',
            'package/README.md', 'package/LICENSE', 'package/THIRD-PARTY-NOTICES.md', 'CHANGELOG.md') {
            $attributes | Should -Match ('(?m)^' + [regex]::Escape($pattern) + ' -text') -Because $pattern
        }
    }
}

Describe 'No cached assembly' {
    It 'compiles the helper in memory and neither reads nor writes a cached assembly' {
        $module = Get-Content -LiteralPath (Join-Path $script:Source 'MeetingCleanupOnPrem.psm1') -Raw
        $module | Should -Not -Match 'LocalApplicationData|LOCALAPPDATA|OutputAssembly|Add-Type\s+-Path|\.dll'
        $module | Should -Match 'Add-Type -TypeDefinition'
        $before = Get-ChildItem -Path (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'MeetingCleanupOnPrem') -Filter '*.dll' -ErrorAction SilentlyContinue | ForEach-Object { "$($_.Name)|$($_.LastWriteTimeUtc.Ticks)" }
        $run = Invoke-Child ("Import-Module '$(Join-Path $script:Source 'MeetingCleanupOnPrem.psd1')' 3>`$null; [MeetingCleanupOnPremNative.Fast].Assembly.Location")
        $run.Text.Trim() | Should -BeNullOrEmpty -Because 'an in-memory assembly has no file location'
        $after = Get-ChildItem -Path (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'MeetingCleanupOnPrem') -Filter '*.dll' -ErrorAction SilentlyContinue | ForEach-Object { "$($_.Name)|$($_.LastWriteTimeUtc.Ticks)" }
        @($after) | Should -Be @($before)
    }
}
Describe 'Owner signatures' -Tag 'RealSignature' {
    # MCO_RELEASE_PACKAGE: the folder of the sealed release package (the extracted ZIP). Without it these cases are skipped.
    BeforeAll { $script:Release = $env:MCO_RELEASE_PACKAGE }
    It 'requires valid timestamped owner signatures on the PowerShell files and the manifest' -Skip:(-not $env:MCO_RELEASE_PACKAGE) {
        $names = @(Get-ChildItem -LiteralPath $script:Release -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' } | Where-Object { $_.FullName -notlike '*\config\*' })
        $names.Count | Should -BeGreaterThan 12
        foreach ($path in @($names.FullName) + @((Join-Path $script:Release 'MeetingCleanupOnPrem.psd1'), (Join-Path $script:Release 'PackageIntegrity.psd1'))) {
            $signature = Get-AuthenticodeSignature -LiteralPath $path
            [string]$signature.Status | Should -Be 'Valid' -Because $path
            $signature.SignerCertificate.Thumbprint | Should -BeExactly $script:Signer
            $signature.TimeStamperCertificate | Should -Not -BeNullOrEmpty
        }
    }
    It 'starts the entry script of the sealed package up to the configuration, and refuses a copy with one changed template byte' -Skip:(-not $env:MCO_RELEASE_PACKAGE) {
        $run = Invoke-Child ("& '$(Join-Path $script:Release 'Invoke-MeetingCleanupOnPrem.ps1')' -ConfigPath '$($script:NoConfig)'; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'no-such-config'
        $run.Text | Should -Not -Match 'Package integrity|positional'
        $copy = Join-Path $TestDrive 'real-entry'
        Copy-Item -LiteralPath $script:Release -Destination $copy -Recurse
        [IO.File]::AppendAllBytes((Join-Path $copy 'templates\Report.template.html'), [byte[]](32))
        $changed = Invoke-Child ("& '$(Join-Path $copy 'Invoke-MeetingCleanupOnPrem.ps1')' -ConfigPath '$($script:NoConfig)'; 'EXIT=' + `$LASTEXITCODE")
        $changed.Text | Should -Match 'SHA256 mismatch: templates\\Report.template.html'
    }
    It 'rejects a cryptographically valid manifest with its timestamp removed' -Skip:(-not $env:MCO_RELEASE_PACKAGE) {
        Add-Type -AssemblyName System.Security.Cryptography.Pkcs
        $fixture = Join-Path $TestDrive 'real-timestamp'
        Copy-Item -LiteralPath $script:Release -Destination $fixture -Recurse
        $path = Join-Path $fixture 'PackageIntegrity.psd1'
        $encoding = [Text.UTF8Encoding]::new($false, $true)
        $text = $encoding.GetString([IO.File]::ReadAllBytes($path))
        $match = [regex]::Match($text, '(?ms)^# SIG # Begin signature block\r?\n(?<Data>.*?)^# SIG # End signature block\r?\n?\z')
        $match.Success | Should -BeTrue
        $cms = [Security.Cryptography.Pkcs.SignedCms]::new()
        $cms.Decode([Convert]::FromBase64String((($match.Groups['Data'].Value -replace '(?m)^#\s?', '') -replace '\s', '')))
        foreach ($attribute in @($cms.SignerInfos[0].UnsignedAttributes)) {
            foreach ($value in $attribute.Values) {
                $cms.SignerInfos[0].RemoveUnsignedAttribute([Security.Cryptography.AsnEncodedData]::new($attribute.Oid, $value.RawData))
            }
        }
        $base64 = [Convert]::ToBase64String($cms.Encode())
        $lines = for ($i = 0; $i -lt $base64.Length; $i += 64) { '# ' + $base64.Substring($i, [Math]::Min(64, $base64.Length - $i)) }
        [IO.File]::WriteAllBytes($path, $encoding.GetBytes($text.Substring(0, $match.Index) +
            "# SIG # Begin signature block`r`n" + ($lines -join "`r`n") + "`r`n# SIG # End signature block`r`n"))
        [string](Get-AuthenticodeSignature -LiteralPath $path).Status | Should -Be 'Valid'
        (Get-AuthenticodeSignature -LiteralPath $path).TimeStamperCertificate | Should -BeNullOrEmpty
        { & (Join-Path $fixture 'src\Test-PackageIntegrity.ps1') -Root $fixture -Distribution } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
    It 'accepts the committed package with the real verifier' -Skip:(-not (Test-Path -LiteralPath (Join-Path (Split-Path $PSScriptRoot -Parent) 'package\PackageIntegrity.psd1'))) {
        $result = & (Join-Path $script:Source 'src\Test-PackageIntegrity.ps1') -Root $script:Source
        $result.Status | Should -Be 'Valid (files only)'
    }
    It 'accepts the release package with the real verifier' -Skip:(-not $env:MCO_RELEASE_PACKAGE) {
        $result = & (Join-Path $script:Release 'src\Test-PackageIntegrity.ps1') -Root $script:Release -Distribution
        $result.Status | Should -Be 'Valid'
    }
    It 'rejects an altered manifest before parsing it, and a copy with one changed byte' -Skip:(-not $env:MCO_RELEASE_PACKAGE) {
        $copy = Join-Path $TestDrive 'real-signature'
        Copy-Item -LiteralPath $script:Release -Destination $copy -Recurse
        [IO.File]::AppendAllBytes((Join-Path $copy 'PackageIntegrity.psd1'), [byte[]](32))
        { & (Join-Path $copy 'src\Test-PackageIntegrity.ps1') -Root $copy -Distribution } | Should -Throw -ExpectedMessage '*owner signature*'
        Remove-Item -LiteralPath $copy -Recurse -Force
        Copy-Item -LiteralPath $script:Release -Destination $copy -Recurse
        [IO.File]::AppendAllBytes((Join-Path $copy 'README.md'), [byte[]](32))
        { & (Join-Path $copy 'src\Test-PackageIntegrity.ps1') -Root $copy -Distribution } | Should -Throw -ExpectedMessage '*SHA256 mismatch: README.md*'
    }
}

AfterAll { Remove-Module MeetingCleanupOnPrem -Force -ErrorAction SilentlyContinue }