# Structured QC rule schema for assays, controls, equipment and errors.
# PowerShell 5.1 friendly: arrays/hashtables only, no classes.

$controlSchemeGroup1 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '1-10'; ExpectedCount = 110 },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '11-20'; ExpectedCount = 100 }
)

$controlSchemeGroup2 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '1-10'; ExpectedCount = 110 },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '11-18'; ExpectedCount = 80 },
    @{ ControlTypeIndex = 2; ControlName = 'Positive Control 2'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '19-20'; ExpectedCount = 20 }
)

$controlSchemeGroup3 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '1-10'; ExpectedCount = 100 },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '11-18'; ExpectedCount = 90 },
    @{ ControlTypeIndex = 2; ControlName = 'Positive Control 2'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '19-20'; ExpectedCount = 20 }
)

$controlSchemeGroup4 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '1-14'; ExpectedCount = 150 },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '15-18'; ExpectedCount = 40 },
    @{ ControlTypeIndex = 2; ControlName = 'Positive Control 2'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '19-20'; ExpectedCount = 20 }
)

$controlSchemeGroup5 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '1-6'; ExpectedCount = 60; ExpectedCategory = 'INVALID' },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '7-18'; ExpectedCount = 130 },
    @{ ControlTypeIndex = 2; ControlName = 'Positive Control 2'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '19-20'; ExpectedCount = 20 }
)

$controlSchemeGroup6 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '1-10'; ExpectedCount = 110 },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '11-16'; ExpectedCount = 60 },
    @{ ControlTypeIndex = 2; ControlName = 'Positive Control 2'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '17-17'; ExpectedCount = 10 },
    @{ ControlTypeIndex = 3; ControlName = 'Positive Control 3'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '18-18'; ExpectedCount = 10 },
    @{ ControlTypeIndex = 4; ControlName = 'Positive Control 4'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '19-19'; ExpectedCount = 10 },
    @{ ControlTypeIndex = 5; ControlName = 'Positive Control 5'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '20-20'; ExpectedCount = 10 }
)

$controlSchemeGroup7 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '1-10'; ExpectedCount = 110 },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '11-17'; ExpectedCount = 70 },
    @{ ControlTypeIndex = 2; ControlName = 'Positive Control 2'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '18-19'; ExpectedCount = 20 },
    @{ ControlTypeIndex = 3; ControlName = 'Positive Control 3'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '20-20'; ExpectedCount = 10 }
)

$controlSchemeGroup8 = @(
    @{ ControlTypeIndex = 0; ControlName = 'Negative Control 1'; ControlLabel = 'NEG'; ExpectedBagRange = '0-10'; ExpectedReplicateRange = '1-10'; ExpectedCount = 110 },
    @{ ControlTypeIndex = 1; ControlName = 'Positive Control 1'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '11-14'; ExpectedCount = 40 },
    @{ ControlTypeIndex = 2; ControlName = 'Positive Control 2'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '15-16'; ExpectedCount = 20 },
    @{ ControlTypeIndex = 3; ControlName = 'Positive Control 3'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '17-18'; ExpectedCount = 20 },
    @{ ControlTypeIndex = 4; ControlName = 'Positive Control 4'; ControlLabel = 'POS'; ExpectedBagRange = '1-10'; ExpectedReplicateRange = '19-20'; ExpectedCount = 20 }
)

$QcSchemaData = @{
    GeneralRules = @{
        SampleIdPattern = @{
            Regex        = '^(?<Material>[A-Za-z0-9\\-]+)_(?<Bag>\\d{1,2})_(?<Index>\\d{1,2})_(?<Replicate>\\d{1,2})(?<Tags>[A-Za-z]*)$'
            Description  = 'PREFIX_BAG_IDX_POS with optional rerun/dilution tags (A/AA/AAA/D##).'
            BagRange     = '00-10'
            IndexRange   = '0-5'
            ReplicateRange = '01-20'
        }
        MaxPressurePSI = 90
        Notes          = 'Derived from GeneralRules sheet in AssayRuleBank_TabBased_updated.xlsx.'
    }
    ControlMaterials = @{
        '001-1688' = @{
            Name         = 'Xpert CARBA-R Negative'
            Role         = 'Control'
            Polarity     = 'Neg'
            Category     = $null
            ActiveInSweden = $true
            Notes        = $null
            Source       = @{ File='ControlMaterialMap_SE.xlsx'; Sheet='PartNoMaster'; Row=6 }
        }
        '001-1686' = @{
            Name         = 'Xpert CARBA-R High POS'
            Role         = 'Control'
            Polarity     = 'Pos'
            Category     = $null
            ActiveInSweden = $true
            Notes        = $null
            Source       = @{ File='ControlMaterialMap_SE.xlsx'; Sheet='PartNoMaster'; Row=7 }
        }
        '001-1687' = @{
            Name         = 'Xpert CARBA-R Low POS'
            Role         = 'Control'
            Polarity     = 'Pos'
            Category     = $null
            ActiveInSweden = $false
            Notes        = 'Not listed in AssayUsage; include for completeness.'
            Source       = @{ File='ControlMaterialMap_SE.xlsx'; Sheet='AssayUsage'; Row=$null }
        }
    }
    ErrorCodes = @(
        @{ Code = '5001'; Name = 'Curve Fit Error'; Group = 'CurveFit'; Classification = 'FUNCTIONAL_MINOR_CURVEFIT'; Description = 'Curve fit error – true functional failure, no re-test'; GeneratesRetest = $false; Source = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='ErrorBank'; Row=2 } },
        @{ Code = '5002'; Name = 'Curve Fit Error'; Group = 'CurveFit'; Classification = 'FUNCTIONAL_MINOR_CURVEFIT'; Description = 'Curve fit error – true functional failure, no re-test'; GeneratesRetest = $false; Source = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='ErrorBank'; Row=3 } },
        @{ Code = '5003'; Name = 'Curve Fit Error'; Group = 'CurveFit'; Classification = 'FUNCTIONAL_MINOR_CURVEFIT'; Description = 'Curve fit error – true functional failure, no re-test'; GeneratesRetest = $false; Source = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='ErrorBank'; Row=4 } },
        @{ Code = '5004'; Name = 'Curve Fit Error'; Group = 'CurveFit'; Classification = 'FUNCTIONAL_MINOR_CURVEFIT'; Description = 'Curve fit error – true functional failure, no re-test'; GeneratesRetest = $false; Source = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='ErrorBank'; Row=5 } }
    )
    Equipment = @{
        Instruments = @()
        Pipettes    = @()
    }
    Assays = @(
        @{
            AssayKey      = 'MTB_RIF'
            AssayName     = 'MTB RIF'
            AssayFamily   = 'MTB'
            AssayVersion  = 'G4'
            AliasPatterns = @('(?i)^MTB\\s*RIF$','(?i)Xpert\\s*MTB[-\\s]*RIF')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'FLUVID_PLUS'
            AssayName     = 'FLUVID+'
            AssayFamily   = 'FLUVID'
            AssayVersion  = 'Plus'
            AliasPatterns = @('(?i)^FLUVID\\s*\\+$','(?i)^FLUVID\\s*PLUS$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'FLUVID'
            AssayName     = 'FLUVID'
            AssayFamily   = 'FLUVID'
            AssayVersion  = 'Standard'
            AliasPatterns = @('(?i)^FLUVID$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'SARS_COV_2_PLUS'
            AssayName     = 'SARS-COV-2 PLUS'
            AssayFamily   = 'SARS_COV_2'
            AssayVersion  = 'Plus'
            AliasPatterns = @('(?i)^SARS[-\\s]*COV[-\\s]*2\\s*PLUS$','(?i)Xpert\\s*Xpress\\s*SARS[-\\s]*CoV[-\\s]*2\\s*Plus')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'SARS_COV_2'
            AssayName     = 'SARS-COV-2'
            AssayFamily   = 'SARS_COV_2'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^SARS[-\\s]*COV[-\\s]*2$','(?i)Xpert\\s*Xpress\\s*SARS[-\\s]*CoV[-\\s]*2')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'CTNG'
            AssayName     = 'CTNG'
            AssayFamily   = 'CTNG'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^CT\\s*/?\\s*NG$','(?i)^CTNG$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'MRSA_SA'
            AssayName     = 'MRSA SA'
            AssayFamily   = 'MRSA'
            AssayVersion  = 'SA'
            AliasPatterns = @('(?i)^MRSA\\s*SA$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'VAN_AB'
            AssayName     = 'VAN AB'
            AssayFamily   = 'VAN'
            AssayVersion  = 'AB'
            AliasPatterns = @('(?i)^VAN\\s*AB$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'GBS'
            AssayName     = 'GBS'
            AssayFamily   = 'GBS'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^GBS$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup1
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'MTB_ULTRA'
            AssayName     = 'MTB ULTRA'
            AssayFamily   = 'MTB'
            AssayVersion  = 'Ultra'
            AliasPatterns = @('(?i)^MTB\\s*ULTRA$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'MTB_XDR'
            AssayName     = 'MTB XDR'
            AssayFamily   = 'MTB'
            AssayVersion  = 'XDR'
            AliasPatterns = @('(?i)^MTB\\s*XDR$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'C_DIFF'
            AssayName     = 'C.DIFF'
            AssayFamily   = 'C_DIFF'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^C\\.?\\s*DIFF$','(?i)^CLOSTRIDIUM\\s*DIFFICILE$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'EBOLA'
            AssayName     = 'EBOLA'
            AssayFamily   = 'EBOLA'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^EBOLA$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'FLU_RSV'
            AssayName     = 'FLU RSV'
            AssayFamily   = 'FLU_RSV'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^FLU\\s*RSV$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'MRSA_NXG'
            AssayName     = 'MRSA NXG'
            AssayFamily   = 'MRSA'
            AssayVersion  = 'NXG'
            AliasPatterns = @('(?i)^MRSA\\s*NXG$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'NORO'
            AssayName     = 'NORO'
            AssayFamily   = 'NORO'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^NORO$','(?i)^NOROVIRUS$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'STREP_A'
            AssayName     = 'STREP A'
            AssayFamily   = 'STREP'
            AssayVersion  = 'A'
            AliasPatterns = @('(?i)^STREP\\s*A$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup2
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HBV_VL'
            AssayName     = 'HBV VL'
            AssayFamily   = 'HBV'
            AssayVersion  = 'VL'
            AliasPatterns = @('(?i)^HBV\\s*VL$','(?i)Hepatitis\\s*B\\s*Viral\\s*Load')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup3
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HCV_VL'
            AssayName     = 'HCV VL'
            AssayFamily   = 'HCV'
            AssayVersion  = 'VL'
            AliasPatterns = @('(?i)^HCV\\s*VL$','(?i)Hepatitis\\s*C\\s*Viral\\s*Load')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup3
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HCV_VL_FS'
            AssayName     = 'HCV VL FS'
            AssayFamily   = 'HCV'
            AssayVersion  = 'VL_FS'
            AliasPatterns = @('(?i)^HCV\\s*VL\\s*FS$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup3
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HIV_VL'
            AssayName     = 'HIV VL'
            AssayFamily   = 'HIV'
            AssayVersion  = 'VL'
            AliasPatterns = @('(?i)^HIV\\s*VL$','(?i)HIV[-\\s]*1\\s*Viral\\s*Load')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup3
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HIV_VL_XC'
            AssayName     = 'HIV VL XC'
            AssayFamily   = 'HIV'
            AssayVersion  = 'VL_XC'
            AliasPatterns = @('(?i)^HIV\\s*VL\\s*XC$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup3
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HIV_QA'
            AssayName     = 'HIV QA'
            AssayFamily   = 'HIV'
            AssayVersion  = 'QA'
            AliasPatterns = @('(?i)^HIV\\s*QA$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup3
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HIV_QA_XC'
            AssayName     = 'HIV QA XC'
            AssayFamily   = 'HIV'
            AssayVersion  = 'QA_XC'
            AliasPatterns = @('(?i)^HIV\\s*QA\\s*XC$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup3
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'CARBA_R'
            AssayName     = 'CARBA R'
            AssayFamily   = 'CARBA_R'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^CARBA\\s*R$')
            PlatformKits  = @()
            ControlMaterials = @('001-1688','001-1686','001-1687')
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup4
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'HPV'
            AssayName     = 'HPV'
            AssayFamily   = 'HPV'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^HPV$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @(
                @{ TestTypePattern='^Negative Control 1$'; LotPattern='^\\d{5}$'; Idx=0; ExpectedResultRegex='^INVALID$'; ExpectedCategory='INVALID'; Source=@{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='ResultRules'; Row=$null } }
            )
            ControlScheme = $controlSchemeGroup5
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'MTB_JP'
            AssayName     = 'MTB JP'
            AssayFamily   = 'MTB'
            AssayVersion  = 'JP'
            AliasPatterns = @('(?i)^MTB\\s*JP$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup6
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'C_DIFF_JP'
            AssayName     = 'C.DIFF JP'
            AssayFamily   = 'C_DIFF'
            AssayVersion  = 'JP'
            AliasPatterns = @('(?i)^C\\.?\\s*DIFF\\s*JP$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup7
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'RESPIRATORY_PANEL'
            AssayName     = 'RESPIRATORY PANEL'
            AssayFamily   = 'RESPIRATORY_PANEL'
            AssayVersion  = 'Base'
            AliasPatterns = @('(?i)^RESPIRATORY\\s*PANEL$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup8
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'RESPIRATORY_PANEL_JP'
            AssayName     = 'RESPIRATORY PANEL JP'
            AssayFamily   = 'RESPIRATORY_PANEL'
            AssayVersion  = 'JP'
            AliasPatterns = @('(?i)^RESPIRATORY\\s*PANEL\\s*JP$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup8
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        },
        @{
            AssayKey      = 'RESPIRATORY_PANEL_R'
            AssayName     = 'RESPIRATORY PANEL R'
            AssayFamily   = 'RESPIRATORY_PANEL'
            AssayVersion  = 'R'
            AliasPatterns = @('(?i)^RESPIRATORY\\s*PANEL\\s*R$')
            PlatformKits  = @()
            ControlMaterials = @()
            ResultRules   = @()
            ControlScheme = $controlSchemeGroup8
            Source        = @{ File='AssayRuleBank_TabBased_updated.xlsx'; Sheet='AssayMap'; Row=$null }
        }
    )
}

$Global:QcSchema = $QcSchemaData
