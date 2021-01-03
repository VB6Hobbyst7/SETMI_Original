Module Functions

#Region "Constants"

    Public Const SigmaPT As Single = 1.26 ' Priestly Taylor Sigma Used in TSEB <-- Kustas & Norman 1999 used 2.0 JBB, can be overwritten in the cover input
    Public Const Rcanopy As Single = 50 's/m added by JBB for PM following Colaizzi et al 2014, can be overwritten in the cover input
    Public Const RcMax As Single = 1000 's/m added by JBB for PM following Colaizzi et al 2014, can be overwritten by the cover input
    Public Const FcExp As Single = 0.7 'following Choudhury 1994 as cited by Li et al 2005.
    Public Const Sigma As Single = 0.0000000567 ' Stefan-Boltzmann W/m2/k4 JBB
    'Public Const Esoil As Single = 0.97 'Commented out by JBB
    Public Esoil As Single = 0.96 'Added by JBB
    'Public Const Eveg As Single = 0.985 'Veg Emissivity JBB, commented out by JBB
    Public Eveg As Single = 0.98 'Added by JBB
    Public Const Gravity As Single = 9.81 ' Gravity Constant JBB
    'Public Const K As Single = 0.41 'von Karman Constant JBB
    Public Const K As Single = 0.4 'von Karman Constant modified by JBB

    'Public Const LAIMin As Single = 0.11 'Used as a ground cover check in resistance calcs JBB
    'Public Const FcoverMin As Single = 0.0631 'Used in Fc Calcs 'Used as a ground cover check in resistance calcs JBB
    Public Const LAIMin As Single = 0.1 'modified by JBB
    Public Const FcoverMin As Single = 0.063 'modified by JBB, this is for the energy balance
    Public Const FcoverMax As Single = 0.95 'added by JBB, this is for the energy balance

    Public Const NDVImax As Single = 0.9 ' Used in NDVI Calc. & in Fc Calcs ' Geli used 0.925 in his disertation JBB, can be overwritten in the cover input
    Public Const NDVImin As Single = 0.2 ' Used in Fc Calcs 'Used in Geli's Dissertation JBB 'CN SAID THIS SHOULD BE LIKE 0.2, can be overwritten in the cover input

    Public Const SAVImax As Single = 0.8 ' Used in SAVI Calc. JBB
    Public Const SAVImin As Single = 0.1 ' Doesn't seem to be used anywhere JBB, was 0.10 JBB Changed to 0.15 for Fc calcs, can be overwritten in the cover input
    Public Const SAVImaxFc As Single = 0.68 'Used in Fc Calcs in WB, can be overwritten in the cover input
    Public Const SAVIFcExponent As Single = 1 'Used in Fc Calcs in WB following Chaudhury Eq 15 this should be between 1/1.4 and 1/0.8 for planofile to erect leaves, respectively, can be overwritten in the cover input
    Public Const TcIterCnt As Integer = 100 'added by JBB for Tcanopy Iteration following Colaizzi et al 2015 IA Symp Paper
#End Region

#Region "Enumerations" ' Enumerations are lists with integers attached i.e. Default Cover = 0, Water = 1, Ag. = 2, etc.

    Enum Cover
        DefaultCover
        Water
        Agriculture
        Alfalfa
        Wheat
        Arrowweed
        Arundo
        Bare_Soil
        Barren
        Cheat_Grass_and_Other_Weeds
        Conifer
        Corn
        Cotton
        Cottonwood
        Dead_Tamarisk
        Desert_Shrubs
        Dryland_Cotton
        Grass
        Mesophytes
        Mesquite
        Sand_and_Gravel
        Soybean
        Tamarisk
        Unclassified
        Upland_Bushes
        Upland_Vegetation
        Vegetated_Decadent
        Citrus
        Guava
        Coconut
        Vineyards
        Watermelon
        Pineapple
        Banana
        Graviola
        Sapoti
        Acerola
        Sabia
        Mango
        Cashew_Giant
        Cashew_Early
        Mata
        Papaya
        Lemon
        Tangerine
        Orange
        Passion_Fruit
        Banana_Apple
        Mixed_Forest
        Coniferous_Forest
        Transitional_Forest   ' Transitional Forest Savana Amazon with constant height and LAI
        Agro_Forestry_Areas     'Agro_Forestry_Areas
        Broad_Leaved_Forest
        Pasture
        Green_Urban_Areas
        Non_Irrigated_Arable_Land
        Natural_Grassland
        Transitional_Woodland_Shrub
        Sparsely_vegetated_areas
        Eastern_Red_Cedar 'Added by JBB for Micheal Neale's Halsey Project
        Sugar_Cane 'Added by JBB for Raoni in Brazil
        CornNLKcb 'added by JBB for non-linear Kcb
        SoybeanNLKcb 'added by JBB for non-linear Kcb
    End Enum

    Enum ImageSource
        Airborne
        Landsat
        MODIS
        'GOES
        'AVHRR
        Unmanned_Aircraft 'Added by JBB 1/25/2018
    End Enum

    Enum EnergyBalance
        One_Layer
        Two_Source
        'SEBAL_Idaho
        'Existing_Point
        'Existing_Grid
    End Enum

    Enum WaterBalance
        Crop_Coefficient_Point
        Crop_Coefficient_Grid
    End Enum

    Enum DataAssimilation
        No_Assimilation
        Single_Weight
        'Time_Varying_Weight
        'Nudging
    End Enum

    Enum Surface
        Canopy
        Soil
    End Enum

    Enum Stability
        Stable
        Neutral
        Unstable
    End Enum

    Enum IrrigationMethod
        Basin
        Border
        Drip
        Drip_Shaded 'added by JBB per ASCE MOP 70 2ed., 2016, p. 239
        Furrow_Narrow_Bed
        Furrow_Narrow_Bed_Alternating
        Furrow_Wide_Bed
        Furrow_Wide_Bed_Alternating
        Not_Irrigated
        Sprinkler
        Subsurface
    End Enum

    Enum ETExtrapolation
        Evaporative_Fraction
        Reference_Evapotranspiration
    End Enum

    Enum ETReferenceType 'added by JBB
        Short_Grass
        Tall_Alfalfa
    End Enum

    Enum EffectivePrecipType 'added by JBB
        SCS_Curve_Number
        'SCS_Curve_Number_Varying
        Percent_Effective
    End Enum

    Enum TSMInitialTemperature 'added by JBB
        Priestly_Taylor
        Penman_Monteith
    End Enum

    Enum SoilHeatFlux 'added by JBB
        Constant_Ratio
        'Phase_Shifted
    End Enum

    Enum WindAdjustment 'added by JBB
        With_Stability_Terms
        Without_Stability_Terms
        No_Adjustment
    End Enum

    Enum KcbType 'added by JBB
        Fitted_Curve
        VI_Interpolation
        VI_Log_Regression
    End Enum
    Enum FcType 'added by JBB this is for WB
        FAO_56
        Vegetation_Index
    End Enum

    Enum FgMethod 'added by JBB
        Using_Past_LAI
        Tabular_Only
    End Enum

    Enum HcMethod 'added by JBB
        Using_Past_Hc
        Current_Hc_Only
    End Enum

    Enum FcMethod 'added by JBB
        Using_Past_Fc
        Current_Fc_Only
    End Enum

    Enum FcVI 'added by JBB
        Hardcoded
        NDVI
        SAVI
    End Enum
    Enum KcbVI 'added by JBB
        Hardcoded
        NDVI
        SAVI
    End Enum
    Enum AllKcbVI 'added by JBB
        NDVI
        SAVI
    End Enum
    Enum ClumpingD 'added by JBB
        Input_Canopy_Width
        Input_Height_to_Width_Ratio
    End Enum

    Enum SAVIForecast 'added by JBB
        Day_of_Year
        Growing_Degree_Days
    End Enum

    Enum DPLimit 'added by JBB
        Constant_Limit
        Decaying_Limit
        No_Limit
    End Enum
    Enum WBPntMethod 'added by JBB copied from other code above
        Both_Manual_and_Tabular
        Manually_Selected
        Tabular_Input
    End Enum

    Enum WBTeMethod 'added by JBB copied from other code above
        No_Transpiration
        Include_Transpiration
        Include_Transp_frm_Skn_Lyr_Also
    End Enum

    Enum WBSkinMethod 'added by JBB copied from other code above
        No_Skin_Evaporation
        Include_Skin_Evap
    End Enum

#End Region

#Region "Functions"

    ''' <summary>
    ''' Calculates fraction of vegetation cover 
    ''' </summary>
    ''' <param name="Cover">Land Cover Classification</param>
    ''' <param name="NDVI">Normalized Difference Vegetation Index</param>
    ''' <param name="SAVI">Soil Adjusted Vegetation Index</param>
    ''' <returns>Vegetation fraction of cover</returns>
    ''' <remarks></remarks>

    Function calcFc(ByVal Cover As Cover, ByVal NDVI As Single, ByVal SAVI As Single, ByVal LAI As Single, ByVal Bioproperties As Bioproperties) As Single ' Used in Energy Balance updated by JBB
        Dim Fc As Single = 0
        'Added by JBB for user defined values
        Dim VI As Single = NDVI
        Dim VImax As Single = NDVImax
        Dim VImin As Single = NDVImin
        Dim Fc_Exponent As Single = FcExp

        If Bioproperties.FcVIInput <> FcVI.Hardcoded And Bioproperties.MaxVIInput > 0 And Bioproperties.MinVIInput > 0 And Bioproperties.ExpVIInput > 0 Then 'User supplied the necessary data for a user defined Fc relationship
            'Overwrite default parameters with user supplied values
            If Bioproperties.FcVIInput = FcVI.SAVI Then VI = SAVI 'overwrites VI with SAVI
            VImax = Bioproperties.MaxVIInput
            VImin = Bioproperties.MinVIInput
            Fc_Exponent = Bioproperties.ExpVIInput
        End If

        ' for all cases 
        VI = Limit(VI, VImin, VImax) 'Keeps it between 0 and 1, it is further limited later
        Fc = 1 - ((VImax - VI) / (VImax - VImin)) ^ (Fc_Exponent) ' From Choudhury et al 1994 as cited by Li et al 2005 eq. 1 JBB
        'End added by JBB


        'If NDVI <= NDVImin Then 'for bare soil surfaces
        '    Fc = FcoverMin
        'ElseIf NDVI >= NDVImax Then
        '    Fc = 0.95 'Fc is limited to 0.99 in FAO 56 p. 149 JBB
        'ElseIf NDVI > NDVImin And NDVI < NDVImax Then
        '    'Fc = ((NDVI - NDVImin) / (NDVImax - NDVImin)) ^ 2 'Campbell & Norman 1998 call this NDVI* (p 268) and cite Carleson et al (1995) who set it equal to fc,  a reasonable approximateion when zenith angles are small JBB
        '    'Fc = 1 - ((NDVImax - NDVI) / (NDVImax - NDVImin)) ^ (0.7) ' From Choudhury et al 1994 as cited by Li et al 2005 eq. 1 JBB
        '    Fc = 1 - ((NDVImax - NDVI) / (NDVImax - NDVImin)) ^ (FcExp) ' From Choudhury et al 1994 as cited by Li et al 2005 eq. 1 JBB
        'End If



        'Below was commented out by JBB on 1/31/2018, it was confusing given the more general methodology above

        ''special cases
        'Select Case Cover
        '    Case Functions.Cover.Wheat
        '        Fc = 1.19 * NDVI - 0.16 ' from Jose Gonzalez dissertation (this was told to me by Isidro Campos). JBB
        '    Case Functions.Cover.Arrowweed Or Functions.Cover.Tamarisk Or Functions.Cover.Dead_Tamarisk
        '        'Fc = 1.72 * NDVI - 0.15 'From Nagler 2003 as described in Geli's Disertation JBB
        '        Fc = 1.85 * NDVI - 0.08   'Nagler 2009 ' Is this correct? Nagler et al 2009 Synthesis of ground and remote sensing data for monitoring ecosystem functions in the Colorado River Delta, Mexico cites Nagler et al 2001 for fc = 1.8NDVI-0.08
        'End Select

        'Return Limit(Fc, FcoverMin, 0.95) ' FcoverMin is a global constant, limit of Fc in FAO 56 is 0.99 JBB
        Return Limit(Fc, FcoverMin, FcoverMax) ' FcoverMin is a global constant, limit of Fc in FAO 56 is 0.99 JBB
    End Function

    Function calcFcWB(ByVal Cover As Cover, ByVal NDVI As Single, ByVal SAVI As Single, ByVal VIType As KcbVI, ByVal VImax_Input As Single, ByVal VImin_Input As Single, ByVal Fc_Exponent_Input As Single, ByVal FcUpperLimit As Single, ByVal FcLowerLimit As Single) As Single ' Used in Water Balance added by JBB mimicking code above 1/31/2018
        'KcbVI type is used here, because the veg index should be the same for Fc and Kcb in the water balance, there is a msg box warning if they are different we only interpolate one, not both so we use whatever is interpolated
        Dim Fc As Single = 0
        'Added by JBB for user defined values
        Dim VI As Single = SAVI
        Dim VImax As Single = SAVImaxFc
        Dim VImin As Single = SAVImin
        Dim Fc_Exponent As Single = SAVIFcExponent

        If VIType <> KcbVI.Hardcoded And VImax_Input > 0 And VImin_Input > 0 And Fc_Exponent_Input > 0 Then 'User supplied the necessary data for a user defined Fc relationship
            'Overwrite default parameters with user supplied values
            If VIType = KcbVI.NDVI Then VI = NDVI 'overwrites VI with SAVI
            VImax = VImax_Input
            VImin = VImin_Input
            Fc_Exponent = Fc_Exponent_Input
        End If

        ' for all cases 
        VI = Limit(VI, VImin, VImax) 'Keeps it between 0 and 1, it is further limited later
        Fc = 1 - ((VImax - VI) / (VImax - VImin)) ^ (Fc_Exponent) ' From Choudhury et al 1994 as cited by Li et al 2005 eq. 1 JBB

        Return Limit(Fc, FcLowerLimit, FcUpperLimit)
    End Function


    Function calcFcbasedonLAI(ByVal Cover As Cover, ByVal LAI As Single) As Single 'Currently not referenced in any active code, JBB 1/31/2018
        Dim Fc As Single = 0

        'special cases
        Select Case Cover
            Case Functions.Cover.Wheat 'Where is the Equation, is this a work in progress? JBB

            Case Functions.Cover.Arrowweed Or Functions.Cover.Tamarisk Or Functions.Cover.Dead_Tamarisk 'Where is the Equation, is this a work in progress? JBB

            Case Functions.Cover.Transitional_Forest
                LAI = 6.9 'Where is this from JBB
                Fc = 1 - Math.Exp(-0.5 * LAI) 'Equation 3 of Norman et al 95 (N95) JBB
        End Select

        Return Limit(Fc, FcoverMin, 0.95)
    End Function

    Function calcNDVI(ByVal Red As Single, ByVal NIR As Single) As Single
        Return Math.Min((NIR - Red) / (NIR + Red), NDVImax) 'Correct, NDVImax is a global variable JBB
    End Function

    Function calcNDWI(ByVal MIR1 As Single, ByVal NIR As Single) As Single ' Gao 1996 JBB
        Return (NIR - MIR1) / (NIR + MIR1) 'CorrectJBB
    End Function

    Function calcSAVI(ByVal Red As Single, ByVal NIR As Single) As Single
        Return Math.Min(1.5 * (NIR - Red) / (NIR + Red + 0.5), SAVImax) ' Correct, SAVImax is a global variable JBB
    End Function

    Function calcOSAVI(ByVal Red As Single, ByVal NIR As Single) As Single
        Return 1.16 * (NIR - Red) / (NIR + Red + 0.16) ' Correct JBB
    End Function


    Function calcAlbedo(ByVal Red As Single, ByVal NIR As Single, ByVal Green As Single) As Single ' Used in Energy Balance JBB
        'Function calcAlbedo(ByVal Red As Single, ByVal NIR As Single) As Single ' Used in Energy Balance JBB
        'Return 0.512 * Red + 0.418 * NIR ' Brest and Goward 1987 for vegetation, they say 0.526*Green + 0.474*NIR for non-veg. They say Veg is when NIR/RED >2.0. JBB
        If NIR / Green > 2 Then
            Return 0.526 * Green + 0.362 * NIR + 0.112 * 0.5 * NIR ' Brest and Goward 1987 for vegetation JBB
        Else
            Return 0.526 * Green + 0.474 * NIR   ' Brest and Goward 1987 for non-vegetation JBB
        End If
    End Function

    Function calcZenithClumping(ByVal NDVI As Single, ByVal SunZenith As Single, ByVal ViewZenith As Single, ByVal LAI As Single, ByVal Hc As Single, ByVal Cveg As Single, ByVal Fc As Single, ByVal Cover As Cover, ByRef Bioproperties As Bioproperties, ByRef ClumpDMethod As ClumpingD) As Clumping_Output

        Dim Output As New Clumping_Output

        Dim k_Clump As Single = -2.2 'See Kustas and Norman 2000 Eq. 4 JBB
        'Dim fGap As single = Math.Exp(-Cveg * (LAI / Fc) / Math.Cos(ViewZenith)) ' Li et al 2005 Eq 3, Cveg is related to canopy leaf angle, 0.5 is for spherical leaf angle distribution JBB commented out by JBB
        Dim fGap As Single = Math.Exp(-Cveg * (LAI / Fc) / Math.Cos(ToRadians(ViewZenith))) ' Li et al 2005 Eq 3, Cveg is related to canopy leaf angle, 0.5 is for spherical leaf angle distribution JBB added by JBB
        Dim p_Clump As Single = 0
        Dim D As Single = 0
        Dim FlgD As Boolean = False 'added by JBB
        If ClumpDMethod = ClumpingD.Input_Height_to_Width_Ratio Then
            D = Bioproperties.ClumpD
            If Bioproperties.ClumpD <= 0 Then FlgD = True
        Else 'default is use input Wc
            D = Hc / Bioproperties.Wc ' See Eq 3, Kustas & Norman 1999 Hc is canopy heigh, Wc is canopy width JBB this line was copied from below by JBB
            If Bioproperties.Wc <= 0 Then FlgD = True 'Added by JBB
        End If

        'If Bioproperties.Wc <= 0 Then ' Wc is canopy crown width JBB 
        If FlgD = True Then ' Added by JBB
            'p_Clump = 3.8 - 0.46 * 2
            Output.Clump0 = 1
            Output.ClumpSun = 1
            Output.ClumpView = 1
        Else
            'D = Hc / Bioproperties.Wc ' See Eq 3, Kustas & Norman 1999 Hc is canopy heigh, Wc is canopy width JBB
            p_Clump = 3.8 - 0.46 * D 'or 3.34 'Eq 3, Kustas & Norman 1999 JBB
            ' Output.Clump0 = (Math.Log(1 - Fc + Fc * fGap)) * Math.Cos(ViewZenith) / (-Cveg * LAI) ' By rearranging Eq. 2 from Li et al 2005 'Is this justified? JBB commented out by JBB
            Output.Clump0 = (Math.Log(1 - Fc + Fc * fGap)) * Math.Cos(ToRadians(ViewZenith)) / (-Cveg * LAI) ' By rearranging Eq. 2 from Li et al 2005 'Is this justified? JBB added by jbb
            Output.ClumpView = Output.Clump0 / (Output.Clump0 + (1 - Output.Clump0) * Math.Exp(k_Clump * (ViewZenith * Math.PI / 180) ^ p_Clump)) 'Eq 3, Kustas & Norman 1999 JBB
            Output.ClumpSun = Output.Clump0 / (Output.Clump0 + (1 - Output.Clump0) * Math.Exp(k_Clump * (SunZenith * Math.PI / 180) ^ p_Clump)) 'Eq 3, Kustas & Norman 1999 JBB
        End If

        Return Output
    End Function

    Function calcFtheta(ByVal ClumpView As Single, ByVal LAI As Single, ByVal Cveg As Single, ByVal ViewZenith As Single) As Single
        'Dim Ftheta As Single = Math.Min(1 - Math.Exp(-Cveg * ClumpView * LAI / Math.Cos(ViewZenith * Math.PI / 180)), 0.8) 'Eq. 6 Kustas & Norman 2000 'Where does 0.8 limit come from?
        Dim Ftheta As Single = Math.Min(1 - Math.Exp(-Cveg * ClumpView * LAI / Math.Cos(ViewZenith * Math.PI / 180)), 0.95) 'Eq. 6 Kustas & Norman 2000 'The 0.95 limit is from JBB
        Return Ftheta
        If Ftheta = 0 Then
            Ftheta = Ftheta
        End If
    End Function

    Function calcG(ByVal Rn As Single, ByVal LAI As Single, ByVal Cover As Cover) As Single 'Used in OSM - JBB
        Select Case Cover
            Case Functions.Cover.Tamarisk, Functions.Cover.Dead_Tamarisk
                Return (0.09 - 0.015 * LAI) * Rn  ' for Tamarisck
            Case Else
                Return ((0.3324 + (-0.024 * LAI)) * (0.8155 + (-0.3032 * Math.Log(Math.Abs(LAI))))) * Rn 'WHERE IS THIS FROM? JBB
        End Select
    End Function

    Function calcHcAir(ByVal NDVI As Single, ByVal NDWI As Single, ByVal SAVI As Single, ByVal OSAVI As Single, ByVal Fc As Single, ByVal Cover As Cover) As Single
        Dim Hc As Single = -9999 ' This equation is used if no maximum crop height raster image is supplied JBB
        Select Case Cover
            Case Functions.Cover.Agriculture

            Case Functions.Cover.Alfalfa

            Case Functions.Cover.Corn
                Hc = (1.86 * OSAVI - 0.2) * (1 + 0.000000482 * Math.Exp(17.69 * OSAVI)) ' From Anderson et al 2004 Upscaling Paper JBB
            Case Functions.Cover.Soybean
                Hc = (0.55 * OSAVI - 0.02) * (1 + 0.0000998 * Math.Exp(9.52 * OSAVI)) ' From Anderson et al 2004 Upscaling Paper JBB
            Case Functions.Cover.Tamarisk, Functions.Cover.Dead_Tamarisk, Functions.Cover.Upland_Bushes, Functions.Cover.Upland_Vegetation


            Case Functions.Cover.Cotton, Functions.Cover.Dryland_Cotton

            Case Functions.Cover.Arrowweed

            Case Functions.Cover.Grass, Functions.Cover.Cheat_Grass_and_Other_Weeds

            Case Functions.Cover.Bare_Soil, Functions.Cover.Sand_and_Gravel

        End Select

        Return Hc
    End Function

    Function calcLAIAir(ByVal NDVI As Single, ByVal SAVI As Single, ByVal oSAVI As Single, ByVal Cover As Cover, ByVal Fc As Single) As Single
        Dim LAI As Single = 0

        Select Case Cover
            Case Functions.Cover.Corn, Functions.Cover.Soybean
                LAI = (4 * oSAVI - 0.8) * (1 + 0.00000473 * Math.Exp(15.64 * oSAVI)) ' From Anderson et al 2004 Upscaling Paper JBB
            Case Functions.Cover.Tamarisk, Functions.Cover.Dead_Tamarisk, Functions.Cover.Upland_Bushes, Functions.Cover.Upland_Vegetation
                'lai =  0.84 * 0.5781 * Math.Exp(2.9455 * NDVI)
                LAI = 0.5781 * Math.Exp(2.9455 * NDVI)
            Case Functions.Cover.Cotton, Functions.Cover.Dryland_Cotton
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Arrowweed
                'LAI = -2 * Math.Log(1 - Fc)
                LAI = 0.5781 * Math.Exp(2.9455 * NDVI)
            Case Functions.Cover.Grass, Functions.Cover.Cheat_Grass_and_Other_Weeds
                'LAI = 24 * 0.1
                LAI = -2 * Math.Log(1 - Fc) 'This is rearrangement of Eq. 3 from Norman et al 1995. JBB, copied from CalcLAILandsat by JBB 4/25/2018
            Case Functions.Cover.Bare_Soil, Functions.Cover.Sand_and_Gravel
                LAI = -Math.Log((0.69 - Limit(SAVI, 0.17, 0.68999)) / 0.59) / 0.91
            Case Functions.Cover.Watermelon
                LAI = -2 * Math.Log(1 - Fc) 'This is rearrangement of Eq. 3 from Norman et al 1995. JBB, copied from CalcLAILandsat by JBB 4/25/2018
            Case Else
                LAI = -2 * Math.Log(1 - Fc) 'This is rearrangement of Eq. 3 from Norman et al 1995. JBB, copied from CalcLAILandsat by JBB 4/25/2018
        End Select
        Return LAI
        'Return Limit(LAI, 0.1, 6)'commented out by JBB To be consistent with other code
    End Function

    Function calcHcUAV(ByVal NDVI As Single, ByVal NDWI As Single, ByVal SAVI As Single, ByVal OSAVI As Single, ByVal Fc As Single, ByVal Cover As Cover) As Single
        'Copied from calcHcAir by JBB 3/1/2018
        Dim Hc As Single = -9999 ' This equation is used if no maximum crop height raster image is supplied JBB
        Select Case Cover
            Case Functions.Cover.Corn
                Hc = (1.33 * OSAVI + 0.08739) * (1 + 0.00779 * Math.Exp(7.081 * OSAVI)) ' From M. Maguire Thesis, UNL, 2018, Preliminary Ch. 3, draft provided to JBB on 7/13/2018, added by JBB
        End Select
        Return Hc
    End Function

    Function calcLAIUAV(ByVal NDVI As Single, ByVal SAVI As Single, ByVal oSAVI As Single, ByVal Cover As Cover, ByVal Fc As Single) As Single
        Dim LAI As Single = 0
        Select Case Cover
            Case Functions.Cover.Corn
                LAI = (0.04246 * oSAVI - 0.005283) * (1 + 71.45 * Math.Exp(1.294 * oSAVI)) ' From M. Maguire Thesis, UNL, 2018, Preliminary Ch. 3, draft provided to JBB on 7/13/2018, added by JBB
        End Select
        Return LAI
    End Function


    Function calcHcTable(ByVal cover As Cover, ByVal Fc As Single, ByVal bioproperties As Bioproperties) As Single
        Dim Hc As Single
        Hc = bioproperties.HcMin + Fc * (bioproperties.HcMax - bioproperties.HcMin) 'Where is this equation from JBB
        'Select Case cover
        '    Case Functions.Cover.Guava, Functions.Cover.Acerola, Functions.Cover.Sapoti, Functions.Cover.Graviola, Functions.Cover.Sabia
        '        Hc = 3.0
        '    Case Functions.Cover.Mango, Functions.Cover.Papaya
        '        Hc = 3.0
        '    Case Functions.Cover.Cashew_Early, Functions.Cover.Cashew_Giant
        '        Hc = 5.0
        '    Case Functions.Cover.Banana, Functions.Cover.Banana_Apple
        '        Hc = 4.0
        '    Case Functions.Cover.Orange, Functions.Cover.Citrus, Functions.Cover.Lemon, Functions.Cover.Tangerine
        '        Hc = 4.0
        '    Case Functions.Cover.Mata
        '        Hc = 5.0
        '    Case Functions.Cover.Vineyards, Functions.Cover.Passion_Fruit
        '        Hc = 2.0
        '    Case Functions.Cover.Mixed_Forest
        '        Hc = 4.0
        '    Case Functions.Cover.Coniferous_Forest
        '        Hc = 3
        '    Case Functions.Cover.Transitional_Forest
        '        Hc = 32
        '    Case Functions.Cover.Watermelon
        '        Hc = 0.4
        '    Case Functions.Cover.Pineapple
        '        Hc = 0.6
        '    Case Functions.Cover.Coconut
        '        Hc = 4.0
        '    Case Else
        '        Hc = bioproperties.HcMin + Fc * (bioproperties.HcMax - bioproperties.HcMin)
        'End Select
        Return Hc 'added by JBB
    End Function

    Function calcHcLandSAT(ByVal NDVI As Single, ByVal NDWI As Single, ByVal SAVI As Single, ByVal oSAVI As Single, ByVal Fc As Single, ByVal Cover As Cover) As Single
        Dim Hc As Single = -9999
        Select Case Cover
            Case Functions.Cover.Agriculture
                Dim a As Single = 6.528 : Dim b As Single = -15.29 : Dim q As Single = 0.893 'from alfalfa
                Hc = Math.Log((q / SAVI - 1) / a) / b
            Case Functions.Cover.Alfalfa
                Dim a As Single = 6.528 : Dim b As Single = -15.29 : Dim q As Single = 0.893
                Hc = Math.Log((q / SAVI - 1) / a) / b
            Case Functions.Cover.Corn
                'Hc = (1.2 * NDWI + 0.6) * (1 + 0.04 * Math.Exp(5.3 * NDWI))
                Hc = (1.86 * oSAVI - 0.2) * (1 + 4.82 * (10 ^ -7) * Math.Exp(17.69 * oSAVI)) ' From Anderson et al 2004 Upscaling Paper JBB
            Case Functions.Cover.Soybean
                'Hc = (0.5 * NDWI + 0.26) * (1 + 0.005 * Math.Exp(4.5 * NDWI)) 'based on Landsat
                Hc = (0.55 * oSAVI - 0.02) * (1 + (9.98 * 10 ^ -5) * Math.Exp(9.52 * oSAVI)) 'based on airborne data ' From Anderson et al 2004 Upscaling Paper JBB
            Case Functions.Cover.Cotton, Functions.Cover.Dryland_Cotton
                'Hc = 8.7122 * NDVI ^ 2 + 0.2491 * NDVI
                Hc = Math.Log((0.7879 / SAVI - 1) / 4.949) / -3.582
            Case Functions.Cover.Wheat
                Hc = 0.75 * NDVI - 0.075
        End Select

        Return Hc
    End Function

    Function calcLAILandSAT(ByVal NDVI As Single, ByVal NDWI As Single, ByVal SAVI As Single, ByVal oSAVI As Single, ByVal Cover As Cover, ByVal Fc As Single) As Single
        Dim LAI As Single = -9999

        Select Case Cover
            Case Cover.Agriculture
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Corn, Functions.Cover.Soybean
                'LAI = (2.88 * NDWI + 1.14) * (1 + 0.1039 * Math.Exp(4.1 * NDWI)) '(the 2.615 was 2.88 Anderson et al 2004
                LAI = (4 * oSAVI - 0.8) * (1 + 4.73 * (10 ^ -6) * Math.Exp(15.64 * oSAVI)) 'based on airborne data ' From Anderson et al 2004 Upscaling Paper JBB
            Case Functions.Cover.Tamarisk, Functions.Cover.Dead_Tamarisk, Functions.Cover.Upland_Bushes, Functions.Cover.Upland_Vegetation
                'calcLAI = 0.84 * 0.5781 * Math.Exp(2.9455 * NDVI)
                LAI = 1.0 * 0.5781 * Math.Exp(2.9455 * NDVI)
            Case Functions.Cover.Cotton, Functions.Cover.Dryland_Cotton
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Alfalfa
                LAI = 0.0151 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Arrowweed
                LAI = 0.5781 * Math.Exp(2.9455 * NDVI)
            Case Functions.Cover.Grass, Functions.Cover.Cheat_Grass_and_Other_Weeds
                LAI = -2 * Math.Log(1 - Fc)
                'LAI = 24 * 0.1
            Case Functions.Cover.Bare_Soil, Functions.Cover.Sand_and_Gravel, Functions.Cover.Barren
                LAI = -Math.Log((0.69 - Limit(SAVI, 0.17, 0.68999)) / 0.59) / 0.91
            Case Functions.Cover.Wheat
                If NDVI >= 0.93 Then
                    LAI = 5
                Else
                    LAI = 0.45 * Math.Log(0.76 / (0.93 - NDVI))
                End If
            Case Functions.Cover.Guava, Functions.Cover.Acerola, Functions.Cover.Sapoti, Functions.Cover.Graviola, Functions.Cover.Sabia  'almond function
                LAI = 11.468 * NDVI - 3.2388
            Case Functions.Cover.Mango, Functions.Cover.Cashew_Early, Functions.Cover.Cashew_Giant ' mangrove function
                LAI = 12.74 * NDVI + 1.34
            Case Functions.Cover.Banana, Functions.Cover.Banana_Apple
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Orange, Functions.Cover.Citrus, Functions.Cover.Lemon, Functions.Cover.Tangerine
                LAI = -2 * Math.Log(1 - Fc)
            Case Functions.Cover.Mata
                LAI = 6.91 * NDVI - 2.49
            Case Functions.Cover.Vineyards, Functions.Cover.Passion_Fruit
                LAI = 5.7 * NDVI - 0.25
            Case Functions.Cover.Mixed_Forest
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Coniferous_Forest 'small bushy tree (used almond)
                LAI = 11.468 * NDVI - 3.2388
            Case Functions.Cover.Papaya
                LAI = 11.468 * NDVI - 3.2388
            Case Functions.Cover.Coconut
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Pineapple
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Watermelon
                LAI = -2 * Math.Log(1 - Fc)
            Case Else
                LAI = -2 * Math.Log(1 - Fc) 'This is rearrangement of Eq. 3 from Norman et al 1995. JBB
        End Select

        Return LAI 'Limit(LAI, LAIMin, 8) ' here always minimum LAI = 0.1 for all surfaces
    End Function

    Function calcLAITable(ByVal cover As Cover, ByVal bioproperties As Bioproperties)
        Dim LAI As Single
        LAI = bioproperties.LAITable
        Return LAI
    End Function
    Function calcLAIMODIS(ByVal NDVI As Single, ByVal NDWI As Single, ByVal SAVI As Single, ByVal oSAVI As Single, ByVal Fc As Single, ByVal Cover As Cover) As Single
        Dim LAI As Single = 0

        Select Case Cover
            Case Functions.Cover.Agriculture
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Corn
                'LAI = (2.88 * NDWI + 1.14) * (1 + 0.1039 * Math.Exp(4.1 * NDWI))
                '(the 2.615 was 2.88 Anderson et al 2004
                LAI = (4 * oSAVI - 0.8) * (1 + 4.73 * (10 ^ -6) * Math.Exp(15.64 * oSAVI)) 'based on airborne data' From Anderson et al 2004 Upscaling Paper JBB
            Case Functions.Cover.Soybean
                LAI = (4 * oSAVI - 0.8) * (1 + 4.73 * (10 ^ -6) * Math.Exp(15.64 * oSAVI)) 'based on airborne data' From Anderson et al 2004 Upscaling Paper JBB
                'LAI = 0.0151 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7)) 'from the alfalafa equation
                'LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Tamarisk, Functions.Cover.Dead_Tamarisk, Functions.Cover.Upland_Bushes, Functions.Cover.Upland_Vegetation
                'calcLAI = 0.84 * 0.5781 * Math.Exp(2.9455 * NDVI)
                LAI = 1.0 * 0.5781 * Math.Exp(2.9455 * NDVI)
                'LAI = -2 * Math.Log(1 - Fc)
            Case Functions.Cover.Cotton, Functions.Cover.Dryland_Cotton
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Alfalfa
                LAI = 0.0151 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.85))
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.85)) 'from SEBAL
                'LAI = (0.04591 * SAVI - 0.01458) * (1 + 25.99 * Math.Exp(2.225 * SAVI))
            Case Functions.Cover.Arrowweed
                'LAI = -2 * Math.Log(1 - Fc)
                LAI = 0.5781 * Math.Exp(2.9455 * NDVI)
            Case Functions.Cover.Grass, Functions.Cover.Cheat_Grass_and_Other_Weeds
                'LAI = 24 * 0.1
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
            Case Functions.Cover.Bare_Soil, Functions.Cover.Sand_and_Gravel
                LAI = -Math.Log((0.69 - Limit(SAVI, 0.17, 0.68999)) / 0.59) / 0.91
            Case Functions.Cover.Wheat
                If NDVI >= 0.93 Then
                    LAI = 5
                Else
                    LAI = 0.45 * Math.Log(0.76 / (0.93 - NDVI))
                End If
            Case Functions.Cover.Guava, Functions.Cover.Acerola, Functions.Cover.Sapoti, Functions.Cover.Graviola, Functions.Cover.Sabia  'almond function
                LAI = 11.468 * NDVI - 3.2388
            Case Functions.Cover.Mango, Functions.Cover.Cashew_Early, Functions.Cover.Cashew_Giant ' mangrove function
                LAI = 12.74 * NDVI + 1.34
            Case Functions.Cover.Banana, Functions.Cover.Banana_Apple
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Orange, Functions.Cover.Citrus, Functions.Cover.Lemon, Functions.Cover.Tangerine
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Mata
                LAI = 6.91 * NDVI - 2.49
            Case Functions.Cover.Vineyards, Functions.Cover.Passion_Fruit
                LAI = 5.7 * NDVI - 0.25
            Case Functions.Cover.Mixed_Forest
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Coniferous_Forest 'small bushy tree (used almond)
                LAI = 11.468 * NDVI - 3.2388
            Case Functions.Cover.Papaya
                LAI = 11.468 * NDVI - 3.2388
            Case Functions.Cover.Coconut
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Pineapple
                LAI = 12.5 * NDVI - 1.375
            Case Functions.Cover.Watermelon
                LAI = 0.03399 * Math.Exp(6.218 * Limit(SAVI, 0.17, 0.7))
        End Select

        'Return Limit(LAI, 0.1, 6'Commented out by JBB To be consistent with other code
    End Function

    Function Limit(ByVal Value As Single, ByVal MinimumValue As Single, ByVal MaximumValue As Single) As Single
        Return Math.Min(Math.Max(Value, MinimumValue), MaximumValue) 'Returns the value passed or an upper or lower bound, if such are exceeded
    End Function

    Function calcTaerodynamic(ByVal Tsurface As Single, ByVal Tair As Single, ByVal U As Single, ByVal LAI As Single) As Single
        Return (0.534 * Tsurface) + (0.39 * Tair) + (0.224 * LAI) - (0.192 * U) + 1.67 ' Used in single Source Model
    End Function

    ''' <summary>
    ''' Estimates net radiation for one layer model
    ''' </summary>
    ''' <param name="SAVI">Soil Adjusted Vegetation Index</param>
    ''' <param name="NDVI"></param>
    ''' <param name="Albedo">Cover Reflectance</param>
    ''' <param name="Tair">Air Temperature (K)</param>
    ''' <param name="Temp">Surfce Temperature (K)</param>
    ''' <param name="Rs">Incoming Solar Radiation (W/m^2)</param>
    ''' <param name="Ea">Actual Vapor Pressure (KPa?)</param>
    ''' <param name="Esoil">Soil Emissivity (dimensionless)</param>
    ''' <param name="Eveg">Vegetation Emissivity (dimensionless)</param>
    ''' <param name="Sigma">Stefan-Boltzmann's Constant (W/m^2/K^4)</param>
    ''' <param name="Month">Image Month (integer)</param>
    ''' <param name="Cover">Land Cover Classification</param>
    ''' <returns>Net radiation (W/m^2)</returns>
    ''' <remarks></remarks>

    Function calcRn(ByVal SAVI As Single, ByVal NDVI As Single, ByVal Albedo As Single, ByVal Tair As Single, ByVal Temp As Single, ByVal Rs As Single, _
                    ByVal Ea As Single, ByVal Esoil As Single, ByVal Eveg As Single, ByVal Sigma As Single, ByVal Month As Integer, ByVal Cover As Cover, ByVal Fc As Single) As Single
        Dim clf As Single = 0 'Clear Sky - See Chavez et al 2005 JBB
        'Eveg = 0.98: Esoil = 0.92: clf = 0

        Dim Epss As Single = Fc * Eveg + (1 - Fc) * Esoil 'See Brunsell and Gillies 2002 Eq 6 JBB
        'epsa = 1.24 * (eap * 10 / tap) ^ (1 / 7) 'Clawford-Duchon (1999)'Brutseart 1975 Eq. 11 as cited by Crawford and Duchon 1999 Eq. 16 JBB
        Dim Epsa = clf + (1 - clf) * (1.22 + 0.06 * Math.Sin((Month + 2) * Math.PI / 6)) * (Ea / Tair) ^ (1 / 7) 'Crawford and Duchon 1999 Eq. 20 JBB, Also Eq. 19 of Chavez et al 2005
        Return (1 - Albedo) * Rs + Epsa * Sigma * Tair ^ 4 - Epss * Sigma * Temp ^ 4 ' ***This Equation is Only used for neutral conditions, Why is Temp not Tsoil used? JBB'Chavez et al 2005, J. Hydrometeorology Vol 6, pp. 923-940 Eq. 16
    End Function

    Function calcRaMO_Stable_OLM(ByVal Tair As Single, ByVal Taerodynamic As Single, ByVal U As Single, ByVal Hc As Single, _
    ByVal Rahn As Single, ByVal Ustarn As Single, ByVal ZoM As Single, ByVal Zoh As Single, _
    ByVal d As Single, ByVal K As Single, ByVal Cp As Single, ByVal Z As Single, ByVal Gravity As Single) As Resistances_Output

        Dim Resistances_Output As New Resistances_Output
        Resistances_Output.Rah = Rahn

        'Zom = 0.123 * hc: Zoh = 0.1 * Zom: d = 0.67 * hc
        Dim Rah As Single = Rahn
        Dim Un1 As Single = Ustarn

        For i = 1 To 100
            ' If Abs(tap - Taero) <= 0.001 Then   tap = tap - 0.1

            Dim H As Single = Cp * (Taerodynamic - Tair) / Rah
            Dim L_MO As Single = ((-Un1) ^ 3 * Tair * Cp) / (Gravity * K * H)
            'If (L_MO) > -0.1 And (L_MO) < 0.1 Then L_MO = 100
            'If L_MO < 0# Then Exit For
            Dim s1 As Single = (Z - d) / L_MO
            If s1 > 1 Then
                Resistances_Output.y_m = -5
                Resistances_Output.y_h = -5
            Else
                Resistances_Output.y_m = -5 * s1
                Resistances_Output.y_h = -5 * s1
            End If
            Un1 = (U * K) / (Math.Log((Z - d) / ZoM) - (Resistances_Output.y_m * ((Z - d) / L_MO)) + (Resistances_Output.y_m * (ZoM / L_MO)))
            Dim Rah1 As Single = (Math.Log((Z - d) / Zoh) - (Resistances_Output.y_h * (Z - d) / L_MO) + (Resistances_Output.y_h * (Zoh / L_MO))) / (Un1 * K)
            Rah = Rah1
            'If i = 300 Then MsgBox i
            If Math.Abs(Rah1 - Rah) < 0.001 Then Exit For
        Next i
        'If Abs(L_MO) >= 100 Then
        'rahp = rahn
        'y_m = 0#: y_h = 0#
        'ElseIf L_MO < 0 Then
        'Resistances_stable = Resistances(tap, Taero, u, hc, rahn, ufn1)
        'Else
        ' End If

        Return Resistances_Output ' ***This Equation is Only used in the single source model***
    End Function

    Function calcRaMO_Unstable_OLM(ByVal Tair As Single, ByVal Taerodynamic As Single, ByVal U As Single, ByVal Hc As Single, _
ByVal Rahn As Single, ByVal Ustarn As Single, ByVal ZoM As Single, ByVal Zoh As Single, ByVal d As Single, ByVal K As Single, _
ByVal Cp As Single, ByVal z As Single, ByVal Gravity As Single) As Resistances_Output

        Dim Un1 As Single = Ustarn
        Dim Resistances_Output As New Resistances_Output
        Resistances_Output.Rah = Rahn
        'tap = tap '- 273.15 it should be in Kelvin.
        'Taero = Taero '- 273.15
        For I = 1 To 100
            ' If Math.Abs(Tair - Taerodynamic) = 0 Then Tair = Tair - 0.1 'need
            Dim H As Single = Cp * (Taerodynamic - Tair) / Resistances_Output.Rah
            Dim L_MO As Single = ((-Un1 ^ 3) * Tair * Cp) / (Gravity * K * H)
            'If (L_MO) > -0.1 And (L_MO) < 0.1 Then L_MO = -100
            'If L_MO > 0 Then Exit For
            Dim x As Single = (1 - 16 * ((z - d) / L_MO)) ^ 0.25
            Resistances_Output.y_h = 2 * Math.Log((1 + x ^ 2) / 2)
            Resistances_Output.y_m = (2 * Math.Log((1 + x) / 2)) + Math.Log((1 + x ^ 2) / 2) - (2 * Math.Atan(x)) + (Math.PI / 2)
            Un1 = (U * K) / (Math.Log((z - d) / ZoM) - (Resistances_Output.y_m * ((z - d) / L_MO)) + (Resistances_Output.y_m * (ZoM / L_MO)))
            Dim Rah1 As Single = (Math.Log((z - d) / Zoh) - (Resistances_Output.y_h * (z - d) / L_MO) + (Resistances_Output.y_h * (Zoh / L_MO))) / (Un1 * K)
            Resistances_Output.Rah = Rah1

            If Math.Abs(Rah1 - Resistances_Output.Rah) < 0.001 Then Exit For
        Next I
        'If Abs(L_MO) >= 100 Then
        'rahp = rahn
        'y_m = 0#: y_h = 0#
        'ElseIf L_MO > 0 Then
        'Resistances = Resistances_stable(tap, Taero, u, hc, rahn, ufn1)
        'Else
        'End If
        Return Resistances_Output ' ***This Equation is Only used in the single source model***
    End Function


    'Function calcRaMO_Stable_TSM(ByRef Resistances As Resistances_Output, ByVal Tair As single, ByVal Hc As single, ByVal L_MO As single, ByVal U As single, ByVal Fc As single, ByVal Clump0 As single, ByVal LAI As single, ByRef TInitial As TInitial_Output, ByVal Temp As single, ByVal fTheta As single, ByVal Rs As single, ByVal Ea As single, ByVal Month As single, ByVal Zom As single, ByVal Zoh As single, ByVal D As single, ByVal Z As single, ByVal Zt As single, ByVal Ag As single, ByVal S As single, ByVal RnCoefficients As RnCoefficients_Output, ByVal W As W_Output, ByVal Cover As Cover) As EnergyComponents_Output'Commented out by JBB
    Function calcRaMO_Stable_TSM(ByRef Resistances As Resistances_Output, ByVal Tair As Single, ByVal Hc As Single, ByVal L_MO As Single, ByVal U As Single, ByVal Fc As Single, ByVal Clump0 As Single, ByVal LAI As Single, ByRef TInitial As TInitial_Output, ByVal Temp As Single, ByVal fTheta As Single, ByVal Rs As Single, ByVal Ea As Single, ByVal Month As Single, ByVal Zom As Single, ByVal Zoh As Single, ByVal D As Single, ByVal Z As Single, ByVal Zt As Single, ByVal Ag As Single, ByVal S As Single, ByVal RnCoefficients As RnCoefficients_Output, ByVal W As W_Output, ByVal Cover As Cover, ByVal Bioproperties As Bioproperties, ByVal TiniMethod As Integer, ByVal Esat As Single) As EnergyComponents_Output
        If Cover = Functions.Cover.Watermelon Then
            Cover = Cover
        End If
        Dim Output As New EnergyComponents_Output
        Dim PT As Single = SigmaPT 'Priestly Taylor Coefficient - A global constant JBB
        If Bioproperties.aPTInput > 0 Then PT = Bioproperties.aPTInput 'added by JBB to allow the user to specify aPT in the cover input
        Dim Rcan As Single = Rcanopy 'Canopy Resistance - added by JBB
        If Bioproperties.RcIniInput > 0 Then Rcan = Bioproperties.RcIniInput 'added by JBB to allow the user to specify Initial Rc in the cover input
        Dim RcanMax As Single = RcMax 'Max Canopy Resistance - added by JBB
        If Bioproperties.RcMaxInput > 0 Then RcanMax = Bioproperties.RcMaxInput 'added by JBB to allow the user to specify Max Rc in the cover input
        Dim TcanopyOld As Single = 0 'added by JBB for Tcanopy Iteration follwoing Colaizzi et al 2015 IA Symp Paper

        Resistances = calcResistances(U, Clump0, L_MO, Hc, fTheta, Temp, TInitial.Tsoil, Tair, D, Zom, Z, Zt, Fc, S, Zoh, LAI, W, Cover, Stability.Stable) 'THERE ARE A FEW PROBLEMS WITH THIS CODE JBB

        '**** added by JBB 3/9/2016 because Rn was not carried over from the previous iteration resulting in zero rn, zero Hc initially!
        calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output, LAI, Fc)
        '*****End added by JBB

        'IN THE UNSTABLE CONDITION THE NET RADIATION COMPONENTS ARE CALCULATED HERE, THIS IS UNNECESSARY IN THE STABLE CONDITION BECASUE WE ALWAYS START UNSTABLE. JBB
        Dim W2 As Single
        Dim Delta2 As Single 'added by JBB
        Dim GammaStar2 As Single 'added by JBB

        '****Added by JBB to get the first Hcanopy to have a value***** Using Original Formulations of the initial T eq's instead of substituting in previously calculated Hcanopy gave the Same Results as Including this step!!!!!!!!!!!
        'W2 = calcW2(TInitial.Tcanopy, W.Cp2P, W.Lambda, 1)
        W2 = calcW2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, 1) 'Do avg of Tc and Tair following Coliazzi et al 2015 IA Symp. Paper, but we don't iterate
        Delta2 = calcDelta2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda) 'added by JBB for PM
        GammaStar2 = calcGammaStar2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, Rcan, Resistances.Ra) 'added by JBB for PM
        If TiniMethod = 1 Then 'added by JBB for PM
            Output.Hcanopy = Output.RnCanopy - Bioproperties.fg * (Delta2 * Output.RnCanopy / (Delta2 + GammaStar2) + W.CpRho * (Esat - Ea) / (Resistances.Ra * (Delta2 + GammaStar2))) 'Following Colaizzi et al 2014 Eq. 2 solved for Hc and added by JBB
        Else 'added by JBB for PM
            Output.Hcanopy = Output.RnCanopy * (1 - PT * Bioproperties.fg * W2) 'Eq 14 from Norman et al 1995 JBB
        End If 'added by JBB for PM

        Do ' Iterate through see Norman et al 1995 P 270 JBB
            For ii = 1 To TcIterCnt '*********Added by JBB for Tcanopy Iteration Follwing Colaizzi et al IA Meeting Paper
                'If LAI <= 0.1 Or Fc <= 0.063 Then 'LAI min and FcMin are global variables that could be used here JBB
                If LAI <= LAIMin Or Fc <= FcoverMin Then 'to match the rest JBB
                    Resistances.Rsoil = 0.0
                    Resistances.Rsoil = 1 / (0.0025 + 0.012 * Resistances.Usoil) 'Norman et al 1995 Eq. B.1 but a' is = 0.0025 following Kustas & Norman 1999 p 20 first paragraph. JBB

                Else
                    'If TInitial.Tsoil > TInitial.Tcanopy Then 'commented out by JBB 5/10/2016
                    If TInitial.Tsoil - TInitial.Tcanopy > 1 Then 'Given the formulation of the two equations below, this logic makes more sense added by JBB
                        Resistances.Rsoil = 1 / (0.0025 * (TInitial.Tsoil - TInitial.Tcanopy) ^ (1 / 3) + 0.012 * Resistances.Usoil) 'Equation 5 Kustas and Norman 1999 JBB
                    Else
                        Resistances.Rsoil = 1 / (0.0025 + 0.012 * Resistances.Usoil) 'Norman et al 1995 Eq. B.1 but a' is = 0.0025 following Kustas & Norman 1999 p 20 first paragraph. JBB
                    End If
                End If

                TcanopyOld = TInitial.Tcanopy
                '*********End Added by JBB for Tcanopy Iteration Follwing Colaizzi et al IA Meeting Paper

                ' get updated Tcanopy,Tac, and Tsoil based on the calculated Rsoil, Ra, and Rx
                TInitial = calcTcomponents(Temp, TInitial.Tcanopy, TInitial.Tsoil, Tair, Resistances, fTheta, W, Output, Cover, LAI, Fc) '  TInitial is in K 'JBB 

                '*********Added by JBB for Tcanopy Iteration Follwing Colaizzi et al IA Meeting Paper
                calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output, LAI, Fc)
                W2 = calcW2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, 1) 'Do avg of Tc and Tair following Coliazzi et al 2015 IA Symp. Paper
                Delta2 = calcDelta2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda) 'added by JBB for PM
                GammaStar2 = calcGammaStar2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, Rcan, Resistances.Ra) 'added by JBB for PM
                If TiniMethod = 1 Then 'added by JBB for PM
                    Output.Hcanopy = Output.RnCanopy - Bioproperties.fg * (Delta2 * Output.RnCanopy / (Delta2 + GammaStar2) + W.CpRho * (Esat - Ea) / (Resistances.Ra * (Delta2 + GammaStar2))) 'Following Colaizzi et al 2014 Eq. 2 solved for Hc and added by JBB
                Else 'added by JBB for PM
                    Output.Hcanopy = Output.RnCanopy * (1 - PT * Bioproperties.fg * W2) 'Eq 14 from Norman et al 1995 JBB
                End If 'added by JBB for PM

                If Math.Abs(TInitial.Tcanopy - TcanopyOld) < 0.01 Then Exit For 'added by JBB for Tcanopy Iteration following Colaizzi et al 2015 IA Sypm Paper
            Next ii '*********Added by JBB for Tcanopy Iteration Follwing Colaizzi et al IA Meeting Paper

            '###############Commented out For Tcanopy Iteration following Colaizzi et al IA Meeting Paper, now if only one iteration, it matches how SETMI worked before the iteration was added
            ''******Commented out by JBB following PyTSEB, and some testing in a spreadsheet that showed this has a neglible effect on the final solution '<-- This section is no longer commented out
            ''******this code modification may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
            ''here we are getting Rn and G
            'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output, LAI, Fc)

            ''W2 = calcW2(TInitial.Tcanopy, W.Cp2P, W.Lambda, 1) 'PT Partitioning coefficient, where are eq's in calcW2 from? JBB
            'W2 = calcW2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, 1) 'Do avg of Tc and Tair following Coliazzi et al 2015 IA Symp. Paper, but we don't iterate
            'Delta2 = calcDelta2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda) 'added by JBB for PM
            'GammaStar2 = calcGammaStar2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, Rcan, Resistances.Ra) 'added by JBB for PM
            ''******************************************************************** End PyTSEB
            '###############End commented out for Tcanopy Iteration

            'Output.Ecanopy = PT * W2 * Output.RnCanopy 'PT for output in units of energy (no divide by lambda) JBB Commented out by JBB because no fg
            'If LAI <= LAIMin Or Fc <= FcoverMin Then 'logic added by JBB for bare soil
            '    Output.Ecanopy = 0 'Added by JBB for bare soil
            'Else
            If TiniMethod = 1 Then 'Added by JBB for PM
                Output.Ecanopy = Bioproperties.fg * (Delta2 * Output.RnCanopy / (Delta2 + GammaStar2) + W.CpRho * (Esat - Ea) / (Resistances.Ra * (Delta2 + GammaStar2))) 'Following Colaizzi et al 2014 Eq. 2 added by JBB
            Else 'Added by JBB for PM
                Output.Ecanopy = PT * Bioproperties.fg * W2 * Output.RnCanopy 'Added by JBB to include fg
            End If 'Added by JBB for PM

            ' End If

            Output.Hcanopy = Output.RnCanopy - Output.Ecanopy 'Norman et al 1995 Eq 14 (calculating the PT 1st) JBB, This becomes redundant, but I left it in
            'Output.Ecanopy = Output.RnCanopy - Output.Hcanopy 'JBB noticed this discrepency previous to getting PyTSEB, but confirmed it with PyTSEB
            '******************************************************************** End PyTSEB inspired code

            'If LAI <= 0.1 Or Fc <= 0.063 Then
            If LAI <= LAIMin Or Fc <= FcoverMin Then 'to match the rest JBB
                'Output.Hsoil = (TInitial.Tsoil - Tair) * W.CpRho / Resistances.Ra 'IS IT OK TO ASSUME RA IS ALL THAT MATTERS IF THERE IS NO CANOPY? JBB
                Output.Hsoil = (TInitial.Tsoil - Tair) * W.CpRho / (Resistances.Ra + Resistances.Rsoil) 'Parallel Model, following and email from Bill Kustas 7/6/2016 added by JBB
            Else
                Output.Hsoil = (TInitial.Tsoil - TInitial.Tac) * W.CpRho / Resistances.Rsoil 'SEE EQ A.1 NORMAN ET AL 1995 JBB
            End If
            Output.EVsoil = Output.RnSoil - Output.GTotal - Output.Hsoil 'Norman et al 1995 Eq 16 JBB
            Output.PstTlr = PT 'added by JBB to output PT
            Output.Rcpy = Rcan 'added by JBB to output Rcanopy
            'If Esoil < 0 Then Priestley Talor factor has to be adjusted up 0.1 'Actually adjusts it down 0.1 JBB
            If Output.EVsoil < 0 Then
                If TiniMethod = 1 Then 'Added by JBB for PM
                    If Rcan < RcanMax Then ' because of numerical error, this may to larger than 1000 'Added by JBB for PM
                        Rcan += 10 'Added by JBB for PM
                    ElseIf Rcan >= RcanMax Then 'Added by JBB for PM
                        PT = 0.09 'added by JBB to make logic below work
                        Exit Do 'Added by JBB for PM
                    End If 'Added by JBB for PM
                Else 'Added by JBB for PM
                    If PT > 0.1 Then 'Because of numercial error PT may drop to 0.090...09 added by JBB
                        PT -= 0.01
                    ElseIf PT <= 0.1 Then  '' if still evsoil <0 ,pt = 0 and set bowen ratio to 10
                        Exit Do
                    End If
                End If 'Added by JBB for PM
            Else                    ' ' if Evsoil > 0 great'This means that a solution has been reached see P270 Norman et al 1995 JBB
                Exit Do
            End If
        Loop
        'If Output.EVsoil < 0 And PT < 0.1 Then '' if still evsoil <0 ,pt = 0 and set bowen ratio to 10
        If Output.EVsoil < 0 And PT <= 0.1 Then '' if still evsoil <0 ,pt = 0 and set bowen ratio to 10, without the <= I don't think this can ever enter this logic JBB
            Dim BowenRatioSoil As Single = 10 'based on Kustas using BR of soil =10 'All refereneces to this have been commented out by someone JBB
            'Output.EVsoil = (Output.RnSoil - Output.GTotal) / (1 + BowenRatioSoil) 'based on Kustas using BR of soil =10
            Output.EVsoil = 0 ' based on Martha's code
            Output.Hsoil = Output.RnSoil - Output.GTotal '----------> >>>>>>try to check if we have bare soil here to check Ra and Rs=0
            Dim xnum As Single = (Tair / Resistances.Ra + (Temp / (fTheta * Resistances.Rx)) - (((1 - fTheta) / (fTheta * Resistances.Rx)) * Output.Hsoil * Resistances.Rsoil / W.CpRho) + Output.Hsoil / W.CpRho) 'Norman et al 1995 Eq A15 JBB
            Dim xden As Single = (1 / Resistances.Ra + 1 / Resistances.Rx + (1 - fTheta) / (fTheta * Resistances.Rx)) 'Norman et al 1995 Eq A15 JBB
            Dim TacLin As Single = xnum / xden 'Norman et al 1995 Eq A15 JBB
            Dim Te As Single = TacLin * (1 + Resistances.Rx / Resistances.Ra) - Output.Hsoil * Resistances.Rx / W.CpRho - Tair * Resistances.Rx / Resistances.Ra 'Norman et al 1995 Eq A17 JBB
            Dim DeltaTac As Single = (Temp ^ 4 - (1 - fTheta) * (Output.Hsoil / W.CpRho * Resistances.Rsoil + TacLin) ^ 4 - fTheta * Te ^ 4) / (4 * fTheta * Te ^ 3 * (1 + Resistances.Rx / Resistances.Ra) + 4 * (1 - fTheta) * (Output.Hsoil * Resistances.Rsoil / W.CpRho + TacLin) ^ 3) 'Norman et al 1995 Eq A16 JBB
            TInitial.Tac = TacLin + DeltaTac 'Norman et al 1995 A18 JBB

            If LAI <= LAIMin Or Fc <= FcoverMin Then 'Logic added by JBB for bare soil
                TInitial.Tsoil = Tair + Output.Hsoil * (Resistances.Ra + Resistances.Rsoil) / W.CpRho 'Norman et al 1995 Eq. 15 JBB
            Else 'the normal case
                TInitial.Tsoil = TInitial.Tac + Output.Hsoil * Resistances.Rsoil / W.CpRho 'Norman et al 1995 Eq A19 JBB 
            End If
            TInitial.Tcanopy = calcTcanopyFromTsoil(Temp, fTheta, TInitial.Tsoil)
            If LAI <= LAIMin Or Fc <= FcoverMin Then 'Logic added by JBB for bare soil
                '    Output.Hcanopy = 0
                Output.Hcanopy = W.CpRho * (TInitial.Tcanopy - Tair) / Resistances.Ra 'Parallel formulation, following email from Bill Kustas 7/6/2016.
            Else 'the normal case
                Output.Hcanopy = W.CpRho * (TInitial.Tcanopy - TInitial.Tac) / Resistances.Rx 'Kustas and Norman 1999 Eq A.15 JBB
            End If
            'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output)
            'Output.EVsoil = (Output.RnSoil - Output.GTotal) / (1 + BowenRatioSoil)
            'Output.Hsoil = Output.RnSoil - Output.GTotal - Output.EVsoil

            '******Commented out by JBB when going through PyTSEB, doesn't necessarily follow PyTSEB and some testing in a spreadsheet that showed this has a neglible effect on the final solution
            '******this code modification may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
            'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output) 'Appears to recalculate Rn based on updated Temps JBB Commented out by JBB because it doesn't appreciably help and can create unclosure
            '****************End Code influenced by PyTSEB
            Output.Ecanopy = Output.RnCanopy - Output.Hcanopy 'Eq A.18 Kustas and Norman 1999 JBB
            ' if Hcanopy > Rncanopy means no latent heat flux, adjust Bowen ratio for soil and canopy as boc =6 and bos =10
        End If

            If (Output.Hsoil + Output.Hcanopy > 700 Or Output.Hsoil + Output.Hcanopy < -200) Then 'This must be for a debug stop JBB
                Output.Hsoil = Output.Hsoil
            End If

            If Output.Ecanopy < 0 Then  ' based on Bill's Code (Output.Hcanopy > Output.RnCanopy) which can occur when LEC < 0
                Dim BowenRatioSoil As Single = 10 ' based on Bill's Code 'All refereneces to this have been commented out by someone JBB
                Dim BownRatioCanopy As Single = 6 ' based on Bill's Code 'All refereneces to this have been commented out by someone JBB
                'Output.EVsoil = (Output.RnSoil - Output.GTotal) / (1 + BowenRatioSoil) '''''' based on Bill's Code
                'Output.Ecanopy = Output.RnCanopy / (1 + BownRatioCanopy) ''''' based on Bill's Code
                Output.Ecanopy = 0
                Output.Hcanopy = Output.RnCanopy
                'Output.Hcanopy = Output.RnCanopy - Output.Ecanopy
                'Output.Hsoil = Output.RnSoil - Output.GTotal - Output.EVsoil ' based on Bill's Code
                'Dim Tac As single = (Tair - 273.15) + Resistances.Ra * Output.Hsoil / W.CpRho + Resistances.Ra * Output.Hcanopy / W.CpRho ' based on Bill's Code
                'TInitial.Tcanopy = Output.Hcanopy * Resistances.Rx / W.CpRho + Tac ' based on Bill's Code
                'TInitial.Tsoil = Output.Hsoil * Resistances.Rsoil / W.CpRho + Tac ' based on Bill's Code

                '******Commented out or added by JBB following PyTSEB and some testing in a spreadsheet that showed the rncomponents calc here has a neglible effect on the final solution
                '******this code modification may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
                'TInitial.Tcanopy = (Output.Hcanopy * (Resistances.Ra + Resistances.Rx) / W.CpRho) + Tair 'Back Calculating for Tc now that we have adjsuted Hc,' Who changed this Eq? The Eq two lines up from 'Bill's Code' seems more correct Seems like some combo of parallel and series? JBB
                'TInitial.Tsoil = ((Temp ^ 4 - fTheta * TInitial.Tcanopy ^ 4) / (1 - fTheta)) ^ (1 / 4) 'A rearrangement of Kustas and Norman 1999 Eq A.1 JBB
                'Output.Hsoil = W.CpRho * (TInitial.Tsoil - Tair) / (Resistances.Rsoil + Resistances.Ra) 'This equation is not in the literature, but it appears to make sense from the parallel model JBB
                'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output) ' Re calculate Rn based on new Temps JBB
                Output.Hsoil = Math.Min(Output.RnSoil - Output.GTotal, Output.Hsoil) 'Doing this conserves energy at the soil surface Directly from PyTSEB
                Output.GTotal = Math.Max(Output.RnSoil - Output.Hsoil, Output.GTotal) 'Doing this conserves energy at the soil surface Directly from PyTSEB
                '**************End Code influenced by PyTSEB

                'Output.GTotal = Output.RnSoil - Output.Hsoil
                '''''' check if we need to update Tac

            End If

            'Output.Stability = 2

            Return Output
    End Function

    Function calcRaMO_Neutral_TSM(ByRef Resistances As Resistances_Output, ByVal U As Single, ByVal Z As Single, ByVal Zt As Single, ByVal L_MO As Single, ByVal Temp As Single, ByVal D As Single, ByVal fTheta As Single, ByVal Ag As Single, ByVal Tair As Single, ByVal Tsoil As Single, ByVal Rs As Single, ByVal Ea As Single, ByVal RecordDate As Date, ByVal NDVI As Single, ByVal Albedo As Single, ByVal SAVI As Single, ByVal LAI As Single, ByVal S As Single, ByVal Fc As Single, ByVal Zoh As Single, ByVal Zom As Single, ByVal Hc As Single, ByVal Clump0 As Single, ByVal WOutput As W_Output, ByVal RnCoefficients As RnCoefficients_Output, ByVal Cover As Cover) As EnergyComponents_Output
        'Resistances = calcResistances(U, Clump0, L_MO, Hc, fTheta, Temp, Tsoil, Tair, D, Zom, Z, Zt, Fc, S, Zoh, LAI, Cover, Stability.Neutral)
        '***JBB***The TSM No longer Calls this, someone commented that bit of code out, this could be the reason for what appear to be old references like calcRn vs CalcRnComponents ***JBB***
        Dim Output As New EnergyComponents_Output
        Select Case Cover
            Case Functions.Cover.Bare_Soil, Functions.Cover.Sand_and_Gravel 'This logic statement is only in the Neutral Calcs!
                Output.RnSoil = calcRn(SAVI, NDVI, Albedo, Tair, Tsoil, Rs, Ea, Esoil, Eveg, Sigma, RecordDate.Month, Cover, Fc) 'In all other locations CalcRnComponents not CalcRn is used JBB
                Output.RnCanopy = 0
                Output.GTotal = Ag * Output.RnSoil 'Kustas and Norman 1999 Eq A10, Li et al 2005 Eq A16
                Output.RnTotal = Output.RnSoil 'makes sense JBB
                Output.Hsoil = (Temp - Tair) * WOutput.CpRho / Resistances.Ra 'Eq 1 from Kustas and Norman 1999 (applies basic H equation) JBB
                Output.EVsoil = Output.RnSoil - Output.GTotal - Output.Hsoil 'Kustas and Norman 1999 Eq A.18 JBB
                Output.Ecanopy = 0 ' This is soil so no canopy LE JBB
                Output.Hcanopy = 0 'This is soil so no canopy H JBB
                'tsoil = Temp - 273.15
            Case Else
                Dim Rsoil As Single = 1 / (0.0025 + (0.012 * Resistances.Usoil)) 'Norman et al 1995 Eq. B.1 but a' is = 0.0025 following Kustas & Norman 1999 p 20 first paragraph. JBB
                Dim Rx As Single = 90 / LAI * (S / Resistances.UdoZom) ^ (0.5) 'Norman et al Eq A.8 & following para JBB
                Dim Tinitial = calcTInitial(Temp, Tair, fTheta, LAI, Fc) 'This appears to be an itterative solution with this set of Tinitial.... being the first guess based on Tcanopy being the average of Tair and Trad JBB
                calcRnComponents(Tair, Ea, Tinitial.Tsoil, Tinitial.Tcanopy, Rs, Ag, RecordDate.Month, RnCoefficients, Output, LAI, Fc)
                Dim W As Single = calcW2(Tinitial.Tcanopy, WOutput.Cp2P, WOutput.Lambda, 1) 'PT Partitioning coefficient, where are eq's in calcW2 from? JBB
                Output.Ecanopy = SigmaPT * W * Output.RnCanopy 'PT for output in units of energy (no divide by lambda) JBB
                Output.Hcanopy = Output.RnCanopy - Output.Ecanopy 'Norman et al 1995 Eq 14 (calculating the PT 1st) JBB
                Tinitial.Tcanopy = ((Tair - 273.15) / Resistances.Ra + (Temp - 273.15) / (Rsoil * (1 - fTheta)) + Rx / WOutput.CpRho * Output.Hcanopy * (1 / Resistances.Ra + 1 / Rsoil + 1 / Rx)) / (1 / Resistances.Ra + 1 / Rsoil + fTheta / (Rsoil * (1 - fTheta))) 'Norman et al 1995 Eq A.7 with Eq. 14 pulled out of it and substitute Hc JBB
                Dim TcLin As Single = Tinitial.Tcanopy + 273.15 'Shouldn't all of the calculations be done in K? JBB!!!!!
                Dim Td As Single = TcLin * (1 + Rsoil / Resistances.Ra) - Rx / WOutput.CpRho * Output.Hcanopy * (1 + Rsoil / Rx + Rsoil / Resistances.Ra) - Tair * Rsoil / Resistances.Ra 'Norman et al 1995 Eq. A.12 w/ Eq 14 replaced by Hc JBB
                Dim DeltaTc As Single = (Temp ^ 4 - fTheta * TcLin ^ 4 - (1 - fTheta) * Td ^ 4) / ((1 - fTheta) * 4 * Td ^ 3 * (1 + Rsoil / Resistances.Ra) + fTheta * 4 * TcLin ^ 3) 'Norman et al 1995 Eq. A.11 JBB
                Tinitial.Tcanopy += DeltaTc 'Norman et al 1995 Eq A.14
                Dim Tac As Single = Tinitial.Tcanopy - Output.Hcanopy * Rx / WOutput.CpRho 'Rearrangement of Norman et al 1995 Eq. A.3 JBB
                Tinitial.Tsoil = calcTsoilFromTcanopy(Temp - 273.15, fTheta, Tinitial.Tcanopy) 'The unit stuff seems wierd with the temperaturesJBB
                Output.Hsoil = (Tinitial.Tsoil - Tac) * WOutput.CpRho / Rsoil 'Norman et al 1995 Eq. A.1 JBB
                calcRnComponents(Tair, Ea, Tinitial.Tsoil, Tinitial.Tcanopy, Rs, Ag, RecordDate.Month, RnCoefficients, Output, LAI, Fc)
                W = calcW2(Tinitial.Tcanopy, WOutput.Cp2P, WOutput.Lambda, 1) 'PT Partitioning coefficient, where are eq's in calcW2 from? JBB
                Output.Ecanopy = SigmaPT * W * Output.RnCanopy 'PT for output in units of energy (no divide by lambda) JBB
                Output.Hcanopy = Output.RnCanopy - Output.Ecanopy 'Norman et al 1995 Eq 14 (calculating the PT 1st) JBB
                Output.EVsoil = Output.RnSoil - Output.GTotal - Output.Hsoil 'Norman et al 1995 Eq. 16 JBB
        End Select
        Return Output
    End Function

    'Function calcRaMO_Unstable_TSM(ByRef Resistances As Resistances_Output, ByVal Z As single, ByVal Zt As single, ByVal Clump0 As single, ByVal Ag As single, ByVal Ea As single, ByRef TInitial As TInitial_Output, ByVal Rs As single, ByVal Month As single, ByVal RnCoefficients As RnCoefficients_Output, ByVal Hc As single, ByVal fTheta As single, ByVal Temp As single, ByVal Tair As single, ByVal Fc As single, ByVal S As single, ByVal Zoh As single, ByVal LAI As single, ByVal D As single, ByVal L_MO As single, ByVal U As single, ByVal Zom As single, ByVal W As W_Output, ByVal Cover As Cover, ByVal albedo As single) As EnergyComponents_Output 'Commented out by JBB
    Function calcRaMO_Unstable_TSM(ByRef Resistances As Resistances_Output, ByVal Z As Single, ByVal Zt As Single, ByVal Clump0 As Single, ByVal Ag As Single, ByVal Ea As Single, ByRef TInitial As TInitial_Output, ByVal Rs As Single, ByVal Month As Single, ByVal RnCoefficients As RnCoefficients_Output, ByVal Hc As Single, ByVal fTheta As Single, ByVal Temp As Single, ByVal Tair As Single, ByVal Fc As Single, ByVal S As Single, ByVal Zoh As Single, ByVal LAI As Single, ByVal D As Single, ByVal L_MO As Single, ByVal U As Single, ByVal Zom As Single, ByVal W As W_Output, ByVal Cover As Cover, ByVal albedo As Single, ByVal Bioproperties As Bioproperties, InitialHcanopy As Boolean, ByVal TiniMethod As Integer, ByVal Esat As Single) As EnergyComponents_Output
        ' here we will be getting Rx and Ra.
        Resistances = calcResistances(U, Clump0, L_MO, Hc, fTheta, Temp, TInitial.Tsoil, Tair, D, Zom, Z, Zt, Fc, S, Zoh, LAI, W, Cover, Stability.Unstable)
        Dim Output As New EnergyComponents_Output

        'here we are getting Rn and G
        calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output, LAI, Fc)

        'calculate for the new Rsoil value and check if EVsoil< 0
        Dim PT As Single = SigmaPT
        If Bioproperties.aPTInput > 0 Then PT = Bioproperties.aPTInput 'added by JBB to allow the user to specify aPT in the cover input
        Dim Rcan As Single = Rcanopy 'Canopy Resistance - added by JBB
        If Bioproperties.RcIniInput > 0 Then Rcan = Bioproperties.RcIniInput 'added by JBB to allow the user to specify Initial Rc in the cover input
        Dim RcanMax As Single = RcMax 'Max Canopy Resistance - added by JBB
        If Bioproperties.RcMaxInput > 0 Then RcanMax = Bioproperties.RcMaxInput 'added by JBB to allow the user to specify Max Rc in the cover input
        Dim W2 As Single
        Dim Delta2 As Single ' added by JBB for PM
        Dim GammaStar2 As Single 'added by JBB for PM
        Dim TempDiff As Single = 0
        Dim TcanopyOld As Single = 0 'added by JBB for Tcanopy Iteration follwoing Colaizzi et al 2015 IA Symp Paper

        '****Added by JBB to get the first Hcanopy to have a value***** Using Original Formulations of the initial T eq's instead of substituting in previously calculated Hcanopy gave the Same Results as Including this step!!!!!!!!!!!
        'W2 = calcW2(TInitial.Tcanopy, W.Cp2P, W.Lambda, 1)
        W2 = calcW2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, 1) 'Do avg of Tc and Tair following Coliazzi et al 2015 IA Symp. Paper, but we don't iterate
        Delta2 = calcDelta2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda) 'added by JBB for PM
        GammaStar2 = calcGammaStar2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, Rcan, Resistances.Ra) 'added by JBB for PM

        If TiniMethod = 1 Then
            Output.Hcanopy = Output.RnCanopy - Bioproperties.fg * (Delta2 * Output.RnCanopy / (Delta2 + GammaStar2) + W.CpRho * (Esat - Ea) / (Resistances.Ra * (Delta2 + GammaStar2))) 'Following Colaizzi et al 2014 Eq. 2 solved for Hc but using sign convention of Colaizzi et al 2012 Eq. b1 and added by JBB
        Else
            Output.Hcanopy = Output.RnCanopy * (1 - PT * Bioproperties.fg * W2) 'Eq 14 from Norman et al 1995 JBB
        End If

        Do
            For ii = 1 To TcIterCnt 'added by JBB for Tcanopy iteration following Colaizzi et al IA Meeting Paper
                TempDiff = TInitial.Tsoil - TInitial.Tcanopy
                If TempDiff > 10 Then TempDiff = 10
                If LAI <= LAIMin Or Fc <= FcoverMin Then
                    Resistances.Rsoil = 0.0 'Why is Rsoil = 0 if there is no cover? JBB
                    Resistances.Rsoil = 1 / (0.0025 + 0.012 * Resistances.Usoil) 'Kustas and Norman 1999 Eq. 5 JBB
                Else
                    If TempDiff > 1 Then 'Why is the logic for >1 not >0? JBB
                        Resistances.Rsoil = 1 / (0.0025 * (TempDiff) ^ (1 / 3) + 0.012 * Resistances.Usoil) 'Kustas and Norman 1999 Eq. 5 JBB
                    Else
                        Resistances.Rsoil = 1 / (0.0025 + 0.012 * Resistances.Usoil) 'A simiplification of Kustas and Norman 1999 Eq. 5 to avoid the third root of a negative, JBB
                    End If
                End If

                TcanopyOld = TInitial.Tcanopy
                '*********End Added by JBB for Tcanopy Iteration Follwing Colaizzi et al IA Meeting Paper

                ' get updated Tcanopy,Tac, and Tsoil based on the calculated Rsoil, Ra, and Rx
                TInitial = calcTcomponents(Temp, TInitial.Tcanopy, TInitial.Tsoil, Tair, Resistances, fTheta, W, Output, Cover, LAI, Fc) '  TInitial is in K JBB

                '*********Added by JBB for Tcanopy Iteration Follwing Colaizzi et al IA Meeting Paper
                calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output, LAI, Fc)
                W2 = calcW2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, 1) 'Do avg of Tc and Tair following Coliazzi et al 2015 IA Symp. Paper, but we don't iterate
                Delta2 = calcDelta2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda) 'added by JBB for PM
                GammaStar2 = calcGammaStar2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, Rcan, Resistances.Ra) 'added by JBB for PM
                If TiniMethod = 1 Then 'added by JBB for PM
                    Output.Hcanopy = Output.RnCanopy - Bioproperties.fg * (Delta2 * Output.RnCanopy / (Delta2 + GammaStar2) + W.CpRho * (Esat - Ea) / (Resistances.Ra * (Delta2 + GammaStar2))) 'Following Colaizzi et al 2014 Eq. 2 solved for Hc and added by JBB
                Else 'added by JBB for PM
                    Output.Hcanopy = Output.RnCanopy * (1 - PT * Bioproperties.fg * W2) 'Eq 14 from Norman et al 1995 JBB
                End If 'added by JBB for PM

                If Math.Abs(TInitial.Tcanopy - TcanopyOld) < 0.01 Then Exit For 'added by JBB for Tcanopy Iteration following Colaizzi et al 2015 IA Sypm Paper
            Next ii '*******added by JBB for Tcanopy Iteration Following Colaizzi et al 2015 IA Symp Paper

            '#############Commented out by JBB for Tcanopy iteration following Colaizzi et al 2015 IA Symp Paper, So if only one iteration it is the same as SETMI was before
            ''******Commented out by JBB following PyTSEB, and some testing in a spreadsheet that showed this has a neglible effect on the final solution'<--This section no longer commented out
            ''******this code modification may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
            ''here we are getting Rn and G
            'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output, LAI, Fc) 'This was commented out but JBB UNCOMMENTED it, whcih didn't change results JBB

            ''W2 = calcW2(TInitial.Tcanopy, W.Cp2P, W.Lambda, 1) 'PT Partitioning coefficient, where are eq's in calcW2 from? JBB
            'W2 = calcW2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, 1) 'Do avg of Tc and Tair following Coliazzi et al 2015 IA Symp. Paper, but we don't iterate
            'Delta2 = calcDelta2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda) 'added by JBB for PM
            'GammaStar2 = calcGammaStar2((TInitial.Tcanopy + Tair) / 2, W.Cp2P, W.Lambda, Rcan, Resistances.Ra) 'added by JBB for PM
            '##########End Commented out for Tcanopy Iteration

            'Output.Ecanopy = PT * W2 * Output.RnCanopy 'PT for output in units of energy (no divide by lambda) JBB Why doesn't it have fg? Commented out by JBB
            'If LAI <= LAIMin Or Fc <= FcoverMin Then 'logic added by JBB for bare soil
            '    Output.Ecanopy = 0
            'Else 'The normal case
            If TiniMethod = 1 Then
                Output.Ecanopy = Bioproperties.fg * (Delta2 * Output.RnCanopy / (Delta2 + GammaStar2) + W.CpRho * (Esat - Ea) / (Resistances.Ra * (Delta2 + GammaStar2))) 'Following Colaizzi et al 2014 Eq. 2 added by JBB
            Else
                Output.Ecanopy = PT * Bioproperties.fg * W2 * Output.RnCanopy 'Added by JBB to include fg
            End If

            ' End If
            Output.Hcanopy = Output.RnCanopy - Output.Ecanopy 'Norman et al 1995 Eq 14 (calculating the PT 1st) JBB'This becomes redundant now
            'Output.Ecanopy = Output.RnCanopy - Output.Hcanopy 'JBB noticed this discrepency previous to getting PyTSEB, but confirmed it with PyTSEB
            '******************************************************************** End PyTSEB inspired code

            If LAI <= LAIMin Or Fc <= FcoverMin Then
                'Output.Hsoil = (TInitial.Tsoil - Tair) * W.CpRho / Resistances.Ra 'This is Norman et al 1995 Eq. A.1 for bare soil, there is no canopy air temp Tac JBB
                Output.Hsoil = (TInitial.Tsoil - Tair) * W.CpRho / (Resistances.Ra + Resistances.Rsoil) 'Parallel Model, following and email from Bill Kustas 7/6/2016 added by JBB
            Else
                Output.Hsoil = (TInitial.Tsoil - TInitial.Tac) * W.CpRho / Resistances.Rsoil 'Norman et al 1995 Eq A.1 JBB
            End If
            Output.EVsoil = Output.RnSoil - Output.GTotal - Output.Hsoil 'Norman and Campbell 1995 Eq. 16 JBB
            Output.PstTlr = PT 'added by JBB to output PT
            Output.Rcpy = Rcan 'Added by JBB to output Rcanopy
            'If Esoil < 0 Then Priestley Talor factor has to be adjusted up 0.1
            If Output.EVsoil < 0 Then 'So can we choose a larger PT to start with and then adjust it down with this method? JBB can we choose a smaller increment?
                If TiniMethod = 1 Then 'Added by JBB for PM
                    If Rcan < RcanMax Then ' may go above 1000 because of numerical error  'Added by JBB for PM
                        Rcan += 10 'Added by JBB for PM
                    ElseIf Rcan >= RcanMax Then 'Added by JBB for PM
                        PT = 0.09 'added by JBB to make logic below work
                        Exit Do 'Added by JBB for PM
                    End If 'Added by JBB for PM
                Else 'Added by JBB for PM
                    If PT > 0.1 Then 'may drop below 0.1 becasue of numerical error
                        PT -= 0.01
                    ElseIf PT <= 0.1 Then  '' if still evsoil <0 ,pt = 0 and set bowen ratio to 10
                        Exit Do
                    End If
                End If 'Added by JBB for PM
            Else                    ' ' if Evsoil > 0 great ' This means that a solution has been reached see P270 Norman et al 1995 JBB
                Exit Do
            End If
        Loop
        'If Output.EVsoil < 0 And PT < 0.1 Then '' if still evsoil <0 ,pt = 0 and set bowen ratio to 10
        If Output.EVsoil < 0 And PT <= 0.1 Then '' if still evsoil <0 ,pt = 0 and set bowen ratio to 10, without the <= I don't think this can ever enter this logic JBB
            Dim BowenRatioSoil As Single = 10 'based on Kustas using BR of soil =10
            'Output.EVsoil = (Output.RnSoil - Output.GTotal) / (1 + BowenRatioSoil) 'based on Kustas using BR of soil =10
            Output.EVsoil = 0 ' based on Martha's code
            Output.Hsoil = Output.RnSoil - Output.GTotal '----------> >>>>>>try to check if we have bare soil here to check Ra and Rs=0 'Norman et al 1995 Eq. A.14 Rearranged w/ LE = 0 JBB
            Dim xnum As Single = (Tair / Resistances.Ra + (Temp / (fTheta * Resistances.Rx)) - (((1 - fTheta) / (fTheta * Resistances.Rx)) * Output.Hsoil * Resistances.Rsoil / W.CpRho) + Output.Hsoil / W.CpRho) 'Norman et al. 1995 Eq. A.15 JBB
            Dim xden As Single = (1 / Resistances.Ra + 1 / Resistances.Rx + (1 - fTheta) / (fTheta * Resistances.Rx)) 'Norman et al. 1995 Eq. A.15 JBB
            Dim TacLin As Single = xnum / xden 'Norman et al. 1995 Eq. A.15 JBB
            Dim Te As Single = TacLin * (1 + Resistances.Rx / Resistances.Ra) - Output.Hsoil * Resistances.Rx / W.CpRho - Tair * Resistances.Rx / Resistances.Ra 'Norman et al 1995 Eq. A.17 JBB
            Dim DeltaTac As Single = (Temp ^ 4 - (1 - fTheta) * (Output.Hsoil / W.CpRho * Resistances.Rsoil + TacLin) ^ 4 - fTheta * Te ^ 4) / (4 * fTheta * Te ^ 3 * (1 + Resistances.Rx / Resistances.Ra) + 4 * (1 - fTheta) * (Output.Hsoil * Resistances.Rsoil / W.CpRho + TacLin) ^ 3) 'Norman et al 1995 Eq. A.16 JBB
            TInitial.Tac = TacLin + DeltaTac 'Norman et al Eq A. 18 JBB

            If LAI <= LAIMin Or Fc <= FcoverMin Then 'Logic added by JBB for bare soil
                TInitial.Tsoil = Tair + Output.Hsoil * (Resistances.Ra + Resistances.Rsoil) / W.CpRho 'Norman et al 1995 Eq. 15 JBB
            Else 'the normal case
                TInitial.Tsoil = TInitial.Tac + Output.Hsoil * Resistances.Rsoil / W.CpRho 'Norman et al 1995 Eq. A.19 JBB
            End If
            TInitial.Tcanopy = calcTcanopyFromTsoil(Temp, fTheta, TInitial.Tsoil) 'Following Norman et al 1995 p 288
            If LAI <= LAIMin Or Fc <= FcoverMin Then 'Logic added by JBB for bare soil
                '    Output.Hcanopy = 0
                Output.Hcanopy = W.CpRho * (TInitial.Tcanopy - Tair) / Resistances.Ra 'Parallel formulation, following email from Bill Kustas 7/6/2016.
            Else 'the normal case
                Output.Hcanopy = W.CpRho * (TInitial.Tcanopy - TInitial.Tac) / Resistances.Rx 'Norman et al 1995 Eq. A.2 JBB 
            End If
            'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output)
            'Output.EVsoil = (Output.RnSoil - Output.GTotal) / (1 + BowenRatioSoil)
            'Output.Hsoil = Output.RnSoil - Output.GTotal - Output.EVsoil
            '******Commented out by JBB when going through PyTSEB, doesn't necessarily follow PyTSEB and some testing in a spreadsheet that showed this has a neglible effect on the final solution
            '******this code modification may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
            'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output)
            '******************End Code influenced by PyTSEB
            Output.Ecanopy = Output.RnCanopy - Output.Hcanopy ' Norman et al 1995 Eq. 18 JBB
            ' if Hcanopy > Rncanopy means no latent heat flux, adjust Bowen ratio for soil and canopy as boc =6 and bos =10
        End If

            If Output.Ecanopy < 0 Then  ' based on Bill's Code (Output.Hcanopy > Output.RnCanopy) which can occur when LEC < 0
                Dim BowenRatioSoil As Single = 10 ' based on Bill's Code
                Dim BownRatioCanopy As Single = 6 ' based on Bill's Code
                'Output.EVsoil = (Output.RnSoil - Output.GTotal) / (1 + BowenRatioSoil) '''''' based on Bill's Code
                'Output.Ecanopy = Output.RnCanopy / (1 + BownRatioCanopy) ''''' based on Bill's Code
                Output.Ecanopy = 0
                Output.Hcanopy = Output.RnCanopy 'Rearrangement of Norman et al 1995 Eq. 18 for LEc = 0 JBB
                'Output.Hcanopy = Output.RnCanopy - Output.Ecanopy
                'Output.Hsoil = Output.RnSoil - Output.GTotal - Output.EVsoil ' based on Bill's Code
                'Dim Tac As single = (Tair - 273.15) + Resistances.Ra * Output.Hsoil / W.CpRho + Resistances.Ra * Output.Hcanopy / W.CpRho ' based on Bill's Code
                'TInitial.Tcanopy = Output.Hcanopy * Resistances.Rx / W.CpRho + Tac ' based on Bill's Code
                'TInitial.Tsoil = Output.Hsoil * Resistances.Rsoil / W.CpRho + Tac ' based on Bill's Code
                '******Commented out or added by JBB following PyTSEB and some testing in a spreadsheet that showed the rncomponents calc here has a neglible effect on the final solution
                '******this code modification may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
                'TInitial.Tcanopy = (Output.Hcanopy * (Resistances.Ra + Resistances.Rx) / W.CpRho) + Tair ' Who changed this Eq? The Eq two lines up from 'Bill's Code' seems more correct Seems like some combo of parallel and series? JBB
                'TInitial.Tsoil = ((Temp ^ 4 - fTheta * TInitial.Tcanopy ^ 4) / (1 - fTheta)) ^ (1 / 4) 'A rearrangement of Kustas and Norman 1999 Eq A.1 JBB
                'Output.Hsoil = W.CpRho * (TInitial.Tsoil - Tair) / (Resistances.Rsoil + Resistances.Ra) 'Norman et al 1999 Eq. 15 JBB
                'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output)
                Output.Hsoil = Math.Min(Output.RnSoil - Output.GTotal, Output.Hsoil) 'Doing this conserves energy at the soil surface Directly from PyTSEB
                Output.GTotal = Math.Max(Output.RnSoil - Output.Hsoil, Output.GTotal) 'Doing this conserves energy at the soil surface Directly from PyTSEB
                '**************End Code influenced by PyTSEB


                'Output.GTotal = Output.RnSoil - Output.Hsoil'<---Uncommenting this would be the same as Norman elt al 1995 p 270
                '''''' check if we need to update Tac
            End If

            ''''this is from Bill's code. This condition can be detected when calculating new value for L
            ' for other stable conditions (Tsoil-Tair) < 0
            If TInitial.Tsoil - Tair < 0 Then
                ' calculation for this part is placed in the calcResistance function
                'Resistances = calcResistances(U, Clump0, L_MO, Hc, fTheta, Temp, TInitial.Tsoil, Tair, D, Zom, Z, Zt, Fc, S, Zoh, LAI, W, Cover, Stability.Stable)
                'TInitial.Tcanopy = (Tair - 273.15 + Temp - 273.15) / 2
                'TInitial.Tsoil = calcTsoilFromTcanopy(Temp - 273.15, fTheta, TInitial.Tcanopy)
                'calcRnComponents(Tair, Ea, TInitial.Tsoil, TInitial.Tcanopy, Rs, Ag, Month, RnCoefficients, Output)
                'W2 = calcW2(TInitial.Tcanopy, W.Cp2P, W.Lambda, 2)
                'Output.Ecanopy = PT * W2 * Output.RnCanopy
                'Output.Hcanopy = Output.RnCanopy - Output.Ecanopy
                'TInitial = calcTcomponents(Temp, TInitial.Tcanopy, TInitial.Tsoil, Tair, Resistances, fTheta, W, Output, Cover)
                ''TInitial.Tac = TInitial.Tcanopy - Output.Hcanopy * Resistances.Rx / W.CpRho
                'Output.Hsoil = W.CpRho * (TInitial.Tsoil + 273.15 - TInitial.Tac) / (Resistances.Rsoil)
                'Output.EVsoil = Output.RnSoil - Output.GTotal - Output.Hsoil
                'Output.HTotal = Output.Hsoil + Output.Hcanopy
                'Output.Stability = 1
            End If

            Output.HTotal = Output.Hsoil + Output.Hcanopy 'This is redundent with a calculation in the Main module JBB
            Output.ETotal = Output.EVsoil + Output.Ecanopy 'This is redundent with a calculation in the Main module JBB

            'If Output.HTotal = 0 Then
            '    Output.Hsoil = Output.Hsoil
            'End If
            ' this to check for NaN or similar issues
            'If Not (Output.Hsoil = Output.Hsoil) Then
            '    Output.Hsoil = Output.Hsoil
            'End If
            Return Output
    End Function

    Function calcResistances(ByVal U As Single, ByVal Clump0 As Single, ByVal L_MO As Single, ByVal Hc As Single, ByVal fTheta As Single, ByVal Temp As Single, ByVal Tsoil As Single, ByVal Tair As Single, ByVal D As Single, ByVal Zom As Single, ByVal Z As Single, ByVal Zt As Single, ByVal Fc As Single, ByVal S As Single, ByVal Zoh As Single, ByVal LAI As Single, ByVal W As W_Output, ByVal Cover As Cover, ByVal Stability As Stability) As Resistances_Output
        'JBB had a note stating: "Comment in Resistances about PyTSEB influence" <-- I believe this is done below.
        Dim Output As New Resistances_Output
        Dim cd As Single = 0.2 ' Drag coefficient as defined by Goudriaan 1977, Where does 0.2 come from???
        Dim AA As Single
        Select Case Stability
            Case Functions.Stability.Stable ' at stable conditions, model solve as parallel rex term is kept (Kustas code)
                'Output.y_m = 0 : Output.y_h = 0 'These are Psi_m and Psi_h from Norman et al 1995 Eq. 6 '<---Original formulation in SETMI
                'Output.y_m = -5 * (Z - D) / L_MO 'Added by JBB following Ham (2005) Agronomy Micromet. Monograph, but Ham doesn't have the -D, that is after Brutseart 1982 Eqs 4.53, 4.56, 4.58 <-- should check Bussinger 1988 paper
                'Output.y_h = Output.y_m 'Added by JBB following Ham (2005) Agronomy Micromet. Monograph, but Ham doesn't have the -D, that is after Brutseart 1982

                '****Stability Terms add by JBB following Brutseart 2005 Hydrology and Introduction, this follows PyTSEB, therefore it may be subject
                '****to the general public license agreement of PyTSEB see:https://github.com/hectornieto/pyTSEB
                Output.y_m = CalcY_MorHStable(Z, D, L_MO) 'Follows Brutseart 2005 and PyTSEB
                Output.y_h = CalcY_MorHStable(Zt, D, L_MO) 'Follows Brutseart 2005 and PyTSEB
                '*****End added following PyTSEB*******************************************

                'Output.Ustar = U * K / Math.Log((Z - D) / Zom) 'Campbell & Norman 1998 Eq. 5.1 Commented out by JBB
                Output.Ustar = (U * K) / (Math.Log((Z - D) / Zom) - Output.y_m) 'see Brutseart EQ. 4.34 Added by JBB
                If Output.Ustar < 0 Then Output.Ustar = 0.01 'Added by JBB on 8/17/2016 for consistency with unstable
                If Fc <= FcoverMin Or LAI <= LAIMin Then   '#### for bare soil and water surfaces conditions
                    Dim ZohZom As Single = calcZohZom(Output.Ustar, Zom, W, Tair, Tsoil) '***This is from Martha Anderson, verified in email on 7/4/2016******
                    Output.Rx = 10 ^ 4
                    'Output.Ra = (Math.Log((Z - D) / ZohZom) - Output.y_h) / (Output.Ustar * K) 'Martha's code '<-- This is actually very similar to Eq. 6 in Norman et al 1995, but it is missing the -y_m term.
                    Output.Ra = (Math.Log((Zt - D) / ZohZom) - Output.y_h) / (Output.Ustar * K) 'Martha's code '<-- This is actually very similar to Eq. 6 in Norman et al 1995, but it is missing the -y_m term.
                    Output.Usoil = U * (Math.Log((0.05 - D) / Zom)) / (Math.Log((Z - D) / Zom) - Output.y_m) 'Eq B.2 of Norman et al 1995, NO limits because hcanopy and LAI limits should keep this from getting too close to 1 following email from Bill Kustas 7/6/2016 added by JBB
                    Output.Usoil = Math.Max(Output.Usoil, 0)
                Else
                    AA = 0.28 * (Clump0 * LAI) ^ (2 / 3) * Hc ^ (1 / 3) * S ^ (-1 / 3) 'Kustas and Norman 2000 Eq. 8, JBB UNCOMMENTED BY JBB
                    'another way to get AA , Martha's code based on Goudriaan 1977, page 110
                    'Dim xld As Single = LAI / Hc 'Leaf Area Density as defined in Goudriaan 1997 p. 109 COMMENTED OUT BY JBB
                    'Dim xl As Single = 0.01 'xl actually represents S 'this xl is unused! JBB COMMENTED OUT BY JBB
                    'Dim xlm As Single = Math.Sqrt((4 * S) / (Math.PI * xld)) 'Mixing Length or mean dist between leaves Goudriaan 1997 Eq. 4.45 Note that: S as used in Norman et al 1995 is 4*LeafArea/perimeter COMMENTED OUT BY JBB

                    'Output.Ucanopy = U * (Math.Log((Hc - D) / Zom)) / (Math.Log((Z - D) / Zom)) 'This is like Eq. B.4 in Norman et al 1995, BUT OMMITS PHSI-M!!! May be based on Campbell & Norman 1998 Eq. 5.1 for Z and Hc commented out by JBB
                    'Output.Ucanopy = U * (Math.Log((Hc - D) / Zom) - Output.y_m) / (Math.Log((Z - D) / Zom) - Output.y_m) 'This is Eq. B.4 in Norman et al 1995, added by JBB even though psim = 0
                    Output.Ucanopy = U * (Math.Log((Hc - D) / Zom)) / (Math.Log((Z - D) / Zom) - Output.y_m) 'This is Eq. B.4 in Norman et al 1995, added by JBB even though psim = 0, corrected by JBB 5/5/2015 to get psim out of numerator
                    Output.Ucanopy = Math.Max(Output.Ucanopy, 0.1) 'Limited to be a small possitive number by JBB added by JBB
                    'AA = Math.Sqrt(cd * Clump0 * LAI * Hc / xlm) ' Eq 4.49 in Goudriaan 1977 is aa=sqrt(cd*LAI*Hc/(2*xlm*iw)), where iw is the proportionality factor for mean eddy velocity and local wind speed clump is added because the clumping factor essentially reduces the effectiveness of LAI. JBB COMMENTED OUT BY JBB
                    '********BELOW COMMENTED OUT BY JBB the Hc limit is redundent with the 0.95 limit below and in the unstable condition
                    'If (Hc > 0.5) Then
                    '    Output.Usoil = Output.Ucanopy * Math.Exp(-AA * (1 - 0.05 / Hc)) 'Eq B.2 of Norman et al 1995
                    'Else
                    '    Output.Usoil = Output.Ucanopy * Math.Exp(-AA * 0.9) '***WHERE IS THIS FROM????????******
                    'End If
                    '*****END COMMENTED OUT BY JBB
                    '*****ADDED BY JBB
                    Output.Usoil = Output.Ucanopy * Math.Exp(-AA * (1 - 0.05 / Hc)) 'Eq B.2 of Norman et al 1995, NO limits because hcanopy and LAI limits should keep this from getting too close to 1
                    '*****END ADDED BY JBB

                    'Output.Ra = (Math.Log((Z - D) / Zom) - Output.y_m) * (Math.Log((Z - D) / Zom) - Output.y_h) / (U * K ^ 2) 'Norman et al 1995 Eq. 6, With an assumption that wind and temp are measured at the same height!!!note Zom=Zoh is from Kustas and Norman 1999 commented out by JBB
                    Output.Ra = (Math.Log((Z - D) / Zom) - Output.y_m) * (Math.Log((Zt - D) / Zom) - Output.y_h) / (U * K ^ 2) 'Norman et al 1995 Eq. 6, but with a fixed heat term added by JBB note Zom=Zoh is from Kustas and Norman 1999 
                    Dim fLocal As Single = LAI / Fc 'See Kustas and Norman 2000 P 850 1st column 
                    Dim AALocal As Single
                    AALocal = 0.28 * fLocal ^ (2 / 3) * Hc ^ (1 / 3) * S ^ (-1 / 3) 'This is Eq B.3 From Norman et al 1995 Uncommented by JBB
                    'AALocal = Math.Sqrt(cd * fLocal * Hc / xlm) 'This is from Goudriaan 4.49 Commented out by JBB
                    Output.UdoZom = Output.Ucanopy * Math.Exp(-AALocal * (1 - (D + Zom) / Hc)) ' Norman et al 1995 Eq. A.9, Kustas and Norman 2000 EQ. 9.
                    Output.Rx = 90 / LAI * (S / Output.UdoZom) ^ 0.5 'Norman et al 1995 Eq. A.8 
                End If

            Case Functions.Stability.Unstable '#### for bare soil and water surfaces conditions
                'Dim X As Single = (1 - 16 * (Z - D) / L_MO) ^ 0.25 'Similar to Eq. 4.22a in Goudriaan 1977, but the 16 is a 15 in his equation, this is actually from Brutseart P. 70. JBB
                'Dim Xt As Single = (1 - 16 * (Zt - D) / L_MO) ^ 0.25 ' !!!!!!!!! this is actually from Brutseart P. 70, Zt is the height of temp meas, it is not used in Line 913 as it should be!!!!!!!! JBB
                'Output.y_m = 2 * Math.Log((1 + X) / 2) + Math.Log((1 + X ^ 2) / 2) - 2 * Math.Atan(X) + Math.PI / 2 'Brutseart 1982 Eq. 4.50 JBB
                'Output.y_h = 2 * Math.Log((1 + Xt ^ 2) / 2) 'Brutseart 1982 Eq. 4.51

                '****Stability Terms add by JBB following Brutseart 2005 Hydrology and Introduction, this follows PyTSEB, therefore it may be subject
                '****to the general public license agreement of PyTSEB see:https://github.com/hectornieto/pyTSEB
                Output.y_m = CalcY_MUnstable(Z, D, L_MO) 'Following Brutseart 2005 and PyTSEB
                Output.y_h = CalcY_HUnstable(Zt, D, L_MO) 'Following Brutseart 2015 and PyTSEB
                '*****End added following PyTSEB*******************************************

                Output.Ustar = (U * K) / (Math.Log((Z - D) / Zom) - Output.y_m) 'see Brutseart EQ. 4.34 JBB
                If Output.Ustar < 0 Then Output.Ustar = 0.01

                If LAI <= LAIMin Or Fc <= FcoverMin Then '#### for bare soil and water types of surfaces 
                    Dim ZohZom As Single = calcZohZom(Output.Ustar, Zom, W, Tair, Tsoil)
                    Output.Rx = 10 ^ 4
                    'Output.Ra = 1 / (0.0025 + 0.012 * Output.Usoil)
                    'Output.Ra = Math.Log((Z - D) / ZohZom) - Output.y_h
                    'Output.Ra = (Math.Log((Z - D) / ZohZom) - Output.y_h) / (Output.Ustar * K) 'Martha's code '<-- This is actually very similar to Eq. 6 in Norman et al 1995, but it is missing the -y_m term.
                    Output.Ra = (Math.Log((Zt - D) / ZohZom) - Output.y_h) / (Output.Ustar * K) 'Martha's code '<-- This is actually very similar to Eq. 6 in Norman et al 1995, but it is missing the -y_m term.
                    'Output.Rex = ZohZom / (K * Output.Ustar)
                    'Output.Ra = Output.Ra + Output.Rex 'this is not used in Martha's model disalexi
                    Output.Usoil = U * (Math.Log((0.05 - D) / Zom)) / (Math.Log((Z - D) / Zom) - Output.y_m) 'Eq B.2 of Norman et al 1995, NO limits because hcanopy and LAI limits should keep this from getting too close to 1 following email from Bill Kustas 7/6/2016 added by JBB
                    Output.Usoil = Math.Max(Output.Usoil, 0)
                Else
                    '*****JBB added the commented out logic below as a testable option JBB commented out the Martha based code in favor of N95's normal method 
                    '*****THIS WOULD NEED TO BE ADDED FOR AALOCAL ALSO AND IN STABLE!!!!!!!!!!
                    'If Cover = Functions.Cover.Corn Or Cover = Functions.Cover.Grass Or Cover = Functions.Cover.Natural_Grassland Or Cover = Functions.Cover.Pasture Or Cover = Functions.Cover.Sugar_Cane Or Cover = Functions.Cover.Wheat Then ' long narrow leaves, not a comprehensive list! 'added by JBB
                    '    '******This is how SETMI was programed 'JBB
                    '    'AA = 0.28 * (Clump0 * LAI) ^ (2 / 3) * Hc ^ (1 / 3) * S ^ (-1 / 3)
                    '    'another way to get AA , Martha's code based on Goudriaan 1977, page 110
                    '    Dim xld As Single = LAI / Hc 'Leaf Area Density as defined in Goudriaan 1997 p. 109
                    '    Dim xl As Single = 0.01 'xl actually represents S 'this xl is unused! JBB
                    '    Dim xlm As Single = Math.Sqrt((4 * S) / (Math.PI * xld)) 'Mixing Length or mean dist between leaves Goudriaan 1997 Eq. 4.45 Note that: S as used in Norman et al 1995 is 4*LeafArea/perimeter
                    '    AA = Math.Sqrt(cd * Clump0 * LAI * Hc / xlm) 'Goudriaan 1977 Eq 4.49, w/ iw = 0.5 <-- using iw=0.5 should give a constant of 0.40 in N95 Eq. B3, rather than 0.28, so this isn't consistent***clump is added because the clumping factor essentially reduces the effectiveness of LAI. JBB
                    '    'AA = Math.Sqrt(cd * Clump0 * LAI * Hc / (xlm * 2)) 'Goudriaan 1977 Eq. 4.49 and w/ iw = 1, which seems to match N95 Eq. B3 assumptions, added and commented out by JBB for reference
                    'Else 'Use Square leaf equation from Norman et al 1995
                    AA = 0.28 * (Clump0 * LAI) ^ (2 / 3) * Hc ^ (1 / 3) * S ^ (-1 / 3) 'Norman et al 1995 Eq. B.3 added by JBB This is Goudriaan EQ. 4.44 and 4.49 w/ iw=1 ***clump is added because the clumping factor essentially reduces the effectiveness of LAI Some simple comparisons with eddy covariance fluxes from CSP in 2013 suggested this was a better method
                    'End If
                    '******End JBB added logic
                    'Output.Ucanopy = U * Math.Log((Hc - D) / Zom) / (Math.Log((Z - D) / Zom)) 'This is like Eq. B.4 in Norman et al 1995, BUT OMMITS PHSI-M!!! May be based on Campbell & Norman 1998 Eq. 5.1 for Z and Hc commented out by JBB
                    'Output.Ucanopy = U * (Math.Log((Hc - D) / Zom) - Output.y_m) / (Math.Log((Z - D) / Zom) - Output.y_m) 'This is Eq. B.4 in Norman et al 1995, added by JBB to include psim
                    Output.Ucanopy = U * (Math.Log((Hc - D) / Zom)) / (Math.Log((Z - D) / Zom) - Output.y_m) 'This is Eq. B.4 in Norman et al 1995, added by JBB to include psim corrected 5/5/2016 by JBB to take psim out of numerator
                    Output.Ucanopy = Math.Max(Output.Ucanopy, 0.1) 'Limited to be a small possitive number by JBB added by JBB


                    '********BELOW COMMENTED OUT BY JBB the limits are unnecessary because of Hc and LAI limits earlier
                    'If (Hc > 0.5) Then 
                    '    Output.Usoil = Output.Ucanopy * Math.Min(Math.Exp(-AA * (1 - 0.05 / Hc)), 0.95) 'Eq B.2 of Norman et al 1995, BUT WITH A LIMIT!!!
                    'Else
                    '    Output.Usoil = Output.Ucanopy * Math.Min(Math.Exp(-AA * 0.9), 0.95) '***WHERE IS THIS FROM????????****** Commented out by JBB, this limit seemed large
                    'End If
                    '*****END COMMENTED OUT BY JBB
                    '*****ADDED BY JBB
                    Output.Usoil = Output.Ucanopy * Math.Exp(-AA * (1 - 0.05 / Hc)) 'Eq B.2 of Norman et al 1995, NO limits because hcanopy and LAI limits should keep this from getting too close to 1
                    '*****END ADDED BY JBB




                    'Output.Ra = (Math.Log((Z - D) / Zom) - Output.y_m) * (Math.Log((Z - D) / Zom) - Output.y_h) / (U * K ^ 2) 'Norman et al 1995 Eq. 6, but assumes that temperature is measured at same height as wind speed. JBB note Zom=Zoh is from Kustas and Norman 1999  commented out by JBB
                    Output.Ra = (Math.Log((Z - D) / Zom) - Output.y_m) * (Math.Log((Zt - D) / Zom) - Output.y_h) / (U * K ^ 2) 'Norman et al 1995 Eq. 6, but with the heat term fixed added by note Zom=Zoh is from Kustas and Norman 1999 JBB
                    Dim fLocal As Single = LAI / Fc 'this is the same as Fsun used elsewhere
                    Dim AALocal As Single
                    AALocal = 0.28 * fLocal ^ (2 / 3) * Hc ^ (1 / 3) * S ^ (-1 / 3) 'Uncommented by JBB
                    'AALocal = Math.Sqrt(cd * fLocal * Hc / xlm) '***WHERE IS THIS FROM????????******Commented out by JBB, if included should use logic as the example above
                    Output.UdoZom = Output.Ucanopy * Math.Exp(-AALocal * (1 - (D + Zom) / Hc)) ' Norman et al 1995 Eq. A.9
                    Output.Rx = 90 / LAI * (S / Output.UdoZom) ^ 0.5 'Norman et al 1995 Eq. A.8 
                End If

        End Select

        Return Output
    End Function
    Function calcZohZom(ByVal Ustar As Single, ByVal Zom As Single, ByVal W As W_Output, ByVal Tair As Single, ByVal Tsoil As Single)
        Dim ZohZom As Single = 7.79 '2 'k * 7.79 ' This is overwritten twice below JBB
        Dim z0plus As Single = Ustar * Zom / (1.5 * 10 ^ -5) 'This doesn't matter because ZohZom is overwritten below JBB
        ZohZom = Zom * Math.Exp(-K * (4.31 * z0plus ^ 0.247 - 5)) 'This is overwritten below JBB, Cahill et al 1997 Eq. 29
        'ZohZom = (1 / 8) * 0.1 / 2.0

        'Martha's code
        Dim Tair_Tsoil As Single = (Tair + Tsoil) / 2      'should be in K
        Dim xmuo As Single = 1.8325 * 10 ^ -5
        'Dim xmu = (296.16 + 393.16) / (Tair_Tsoil + 393.16) * (Tair_Tsoil / 296.16) * 1.5 * xmuo 'commented out by JBB
        Dim xmu = (296.16 + 393.16) / (Tair_Tsoil + 393.16) * (Tair_Tsoil / 296.16) ^ 1.5 * xmuo 'Fixed by JBB following Code from an Email from M.Anderson 7/4/2016
        Dim xnu = xmu / W.Rho
        Dim restar = Ustar * Zom / xnu
        ZohZom = Zom * Math.Exp(-K * (4.0 * restar ^ 0.15 - 5))
        Return ZohZom
    End Function

    Function calcRnComponents_1(ByVal Tair As Single, ByVal Ea As Single, ByVal Tsoil As Single, ByVal Tcanopy As Single, ByVal Rs As Single, ByVal Ag As Single, ByVal Month As Single, ByVal RnCoefficients As RnCoefficients_Output, ByRef Output As EnergyComponents_Output) As EnergyComponents_Output
        Dim Clf As Single = 0 '***This function calcRnComponents_1 is not used anywhere***JBB
        Dim Eair As Single = Clf + (1 - Clf) * (1.22 + 0.06 * Math.Sin((Month + 2) * Math.PI / 6)) * (Ea / Tair) ^ (1 / 7) 'Crawford and Duchon 1999 Eq. 20 JBB
        'Dim Eair = 1.24 * (Ea / Tair) ^ (1 / 7)'Brutsaert 1975 Eq 11 cited in Crawford and Duchon 1999 Eq. 16 JBB
        Dim RlSky As Single = Eair * Sigma * Tair ^ 4
        Dim RlSoil As Single = Esoil * Sigma * (Tsoil) ^ 4
        Dim RlCanopy = Eveg * Sigma * (Tcanopy) ^ 4
        Output.RnSoil = RnCoefficients.TauThermal * RlSky + (1 - RnCoefficients.TauThermal) * RlCanopy - RlSoil + RnCoefficients.TauSolar * (1 - RnCoefficients.AlbSoil) * Rs '  combo of Li et al 2005 Eq A19 or Kustas and Norman 2000 Eq 3a and Li et al Eq A17 JBB
        Output.RnCanopy = (1 - RnCoefficients.TauThermal) * (RlSky + RlSoil - 2 * RlCanopy) + (1 - RnCoefficients.TauSolar) * (1 - RnCoefficients.AlbCanopy) * Rs ' combo of Li et al 2005 Eq A20 or Kustas and Norman 2000 Eq 3b and Li et al Eq A18 JBB
        Output.RnTotal = Output.RnCanopy + Output.RnSoil ' See first paragraph on P 849 of Kustas and Norman 2000, Li et al 2005 Eq A11.
        Output.GTotal = Ag * Output.RnSoil 'Kustas and Norman 1999 Eq A10, Li et al 2005 Eq A16
        Return Output
    End Function

    'Function calcRnComponents(ByVal Tair As Single, ByVal Ea As Single, ByVal Tsoil As Single, ByVal Tcanopy As Single, ByVal Rs As Single, ByVal Ag As Single, ByVal Month As Single, ByVal RnCoefficients As RnCoefficients_Output, ByRef Output As EnergyComponents_Output) As EnergyComponents_Output
    Function calcRnComponents(ByVal Tair As Single, ByVal Ea As Single, ByVal Tsoil As Single, ByVal Tcanopy As Single, ByVal Rs As Single, ByVal Ag As Single, ByVal Month As Single, ByVal RnCoefficients As RnCoefficients_Output, ByRef Output As EnergyComponents_Output, ByVal LAI As Single, ByVal FractionCover As Single) As EnergyComponents_Output 'Added by JBB for bare soil
        'c     *      LW: using sky, canopy and soil thermal fluxes and
        'c     *          canopy thermal transmission coefficient.
        'c     *      SW: using SDN components, canopy transmission and
        'c     *          soil reflectivity.

        ' sdn is used to compute fvis, fclear, fnir, difvis, difnir
        Dim Clf As Single = 0
        'Dim Eair As single = Clf + (1 - Clf) * (1.22 + 0.06 * Math.Sin((Month + 2) * Math.PI / 6)) * (Ea / Tair) ^ (1 / 7) 'Crawford and Duchon 1999 Eq. 20 JBB
        Dim Eair As Single = 1.24 * (Ea / Tair) ^ (1 / 7) 'Brutsaert 1975 Eq 11 cited in Crawford and Duchon 1999 Eq. 16 JBB
        Dim RlSky As Single = (Eair * RnCoefficients.fclear + 1 - RnCoefficients.fclear) * Sigma * Tair ^ 4  'it includes  the term fclear based on Martha's code (fclear=1) 'This is another way of presenting Eq 11 from Crawford and Duchon 1999 by setting fclear = 1 - clf 'Colaizzi et al 2012 Eq. A8a JBBJBB
        Dim RlSoil As Single = Esoil * Sigma * (Tsoil) ^ 4 'Colaizzi et al 2012 Eq. A8b JBB
        Dim RlCanopy As Single = Eveg * Sigma * (Tcanopy) ^ 4 'Colaizzi et al 2012 Eq. A8c JBB
        Output.RnSoil = RnCoefficients.TauThermal * RlSky + (1 - RnCoefficients.TauThermal) * RlCanopy - RlSoil + RnCoefficients.TauSolar * (1 - RnCoefficients.AlbSoil) * Rs '  combo of Li et al 2005 Eq A19 or Kustas and Norman 2000 Eq 3a and Li et al Eq A17 JBB
        Output.RnCanopy = (1 - RnCoefficients.TauThermal) * (RlSky + RlSoil - 2 * RlCanopy) + (1 - RnCoefficients.TauSolar) * (1 - RnCoefficients.AlbCanopy) * Rs ' combo of Li et al 2005 Eq A20 or Kustas and Norman 2000 Eq 3b and Li et al Eq A18 JBB
        'If LAI <= LAIMin Or FractionCover <= FcoverMin Then 'added by JBB for bare soil, it overrides previous calculations
        '    Output.RnSoil = RlSky - RlSoil + (1 - RnCoefficients.AlbSoil) * Rs 'Simplified RnSoil Equation if there is no canopy
        '    Output.RnCanopy = 0
        'End If
        Output.RnTotal = Output.RnCanopy + Output.RnSoil ' See first paragraph on P 849 of Kustas and Norman 2000, Li et al 2005 Eq A11.
        Output.GTotal = Ag * Output.RnSoil 'Kustas and Norman 1999 Eq A10, Li et al 2005 Eq A16
        Return Output
    End Function

    Function calcRnCoefficients_1(ByVal Rs As Single, ByVal ClumpSun As Single, ByVal LAI As Single, ByVal Zenith As Single, ByVal Cover As Cover, ByVal Fc As Single, ByRef Bioproperties As Bioproperties) As RnCoefficients_Output
        Dim AlphaSoilVis As Single = 0.15  ' this was 0.15  from Kustas code , it indicated high soil ref
        Dim AlphaSoilNIR As Single = 0.25   ' this was 0.25  from Kustas code , it indicated high soil ref

        ' adjusting for cases with short vegetation 
        If Fc <= 0.0631 Or LAI <= 0.1 Then
            'Bioproperties.AlphaVIS = 0.83
            'Bioproperties.AlphaNIR = 0.57
        End If
        'Dim AlphaVis As single = 0.83, AlphaNIR As single = 0.4, AlphaTIR As single = 0.95
        'Select Case Cover
        '    Case Functions.Cover.Corn, Functions.Cover.Soybean
        '        AlphaVis = 0.88 : AlphaNIR = 0.2 : AlphaTIR = 0.95
        '    Case Functions.Cover.Soybean
        '        AlphaVis = 0.85 : AlphaNIR = 0.15 : AlphaTIR = 0.95
        '    Case Functions.Cover.Tamarisk, Functions.Cover.Grass
        '        AlphaVis = 0.88 : AlphaNIR = 0.55 : AlphaTIR = 0.95
        '    Case Functions.Cover.Dead_Tamarisk, Functions.Cover.Arrowweed
        '        AlphaVis = 0.88 : AlphaNIR = 0.55 : AlphaTIR = 0.95
        '    Case Functions.Cover.Bare_Soil
        '        AlphaVis = 0.83 : AlphaNIR = 0.57 : AlphaTIR = 0.95
        '    Case Functions.Cover.Wheat
        '        AlphaVis = 0.83 : AlphaNIR = 0.4 : AlphaTIR = 0.95
        'End Select

        'If LAI <= 0.1 Then 'Or Fc <= 0.0631 
        '    Bioproperties.AlphaVIS = 0.83 : Bioproperties.AlphaNIR = 0.57 : Bioproperties.AlphaTIR = 0.95
        'End If

        Dim Fsun As Single = ClumpSun * LAI
        Dim Xp As Single = 1.0
        Dim Kbe As Single = Math.Sqrt(Xp ^ 2 + (Math.Tan(Zenith * Math.PI / 180)) ^ 2) / (Xp + 1.774 * (Xp + 1.182) ^ (-0.733))

        'set limits for kbe
        Dim Kd As Single = 0
        Select Case Cover
            Case Functions.Cover.Corn, Functions.Cover.Soybean
                If LAI <= 0.5 Then
                    Kd = 0.9
                ElseIf LAI > 0.5 And LAI <= 2 Then
                    Kd = 0.8
                Else
                    Kd = 0.7
                End If
            Case Functions.Cover.Tamarisk, Functions.Cover.Arrowweed
                If LAI <= 0.5 Then
                    Kd = 0.9
                ElseIf LAI > 0.5 And LAI <= 2 Then
                    Kd = 0.85
                Else
                    Kd = 0.7
                End If
            Case Functions.Cover.Bare_Soil, Functions.Cover.Sand_and_Gravel
                'Kd = 1
                If LAI <= 0.5 Then
                    Kd = 0.9
                ElseIf LAI > 0.5 And LAI <= 2 Then
                    Kd = 0.85
                Else
                    Kd = 0.7
                End If
            Case Else
                If LAI <= 0.5 Then
                    Kd = 0.9
                ElseIf LAI > 0.5 And LAI <= 2 Then
                    Kd = 0.8
                Else
                    Kd = 0.7
                End If
        End Select
        Kd = -0.0683 * Math.Log(Fsun) + 0.804 ' from Martha's code'This overwrites all other Kd's!!!

        Dim RefRohVis As Single = (1 - Math.Sqrt(Bioproperties.AlphaVIS)) / (1 + Math.Sqrt(Bioproperties.AlphaVIS))
        Dim RefRohNIR As Single = (1 - Math.Sqrt(Bioproperties.AlphaNIR)) / (1 + Math.Sqrt(Bioproperties.AlphaNIR))
        Dim RefVis As Single = 2 * Kbe / (1 + Kbe) * RefRohVis
        Dim RefNIR As Single = 2 * Kbe / (1 + Kbe) * RefRohNIR

        Dim AlbCanopyVisDir As Single = (RefVis + (RefVis - AlphaSoilVis) / (RefVis * AlphaSoilVis - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * Kbe * Fsun)) / _
        (1 + RefVis * (RefVis - AlphaSoilVis) / (RefVis * AlphaSoilVis - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * Kbe * Fsun))

        Dim AlbCanopyNIRDir As Single = (RefNIR + (RefNIR - AlphaSoilNIR) / (RefNIR * AlphaSoilNIR - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * Kbe * Fsun)) / _
        (1 + RefNIR * (RefNIR - AlphaSoilNIR) / (RefNIR * AlphaSoilNIR - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * Kbe * Fsun))

        Dim TauVisDir As Single = ((RefVis ^ 2 - 1) * Math.Exp(-Math.Sqrt(Bioproperties.AlphaVIS) * Kbe * Fsun)) / _
        ((RefVis * AlphaSoilVis - 1) + RefVis * (RefVis - AlphaSoilVis) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * Kbe * Fsun))

        Dim TauNIRDir As Single = ((RefNIR ^ 2 - 1) * Math.Exp(-Math.Sqrt(Bioproperties.AlphaNIR) * Kbe * Fsun)) / _
        ((RefNIR * AlphaSoilNIR - 1) + RefNIR * (RefNIR - AlphaSoilNIR) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * Kbe * Fsun))

        RefVis = 2 * Kd / (1 + Kd) * RefRohVis
        RefNIR = 2 * Kd / (1 + Kd) * RefRohNIR

        Dim AlbCanopyVisDiff As Single = (RefVis + (RefVis - AlphaSoilVis) / (RefVis * AlphaSoilVis - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * Kd * Fsun)) / _
        (1 + RefVis * (RefVis - AlphaSoilVis) / (RefVis * AlphaSoilVis - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * Kd * Fsun))

        Dim AlbCanopyNIRDiff As Single = (RefNIR + (RefNIR - AlphaSoilNIR) / (RefNIR * AlphaSoilNIR - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * Kd * Fsun)) / _
        (1 + RefNIR * (RefNIR - AlphaSoilNIR) / (RefNIR * AlphaSoilNIR - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * Kd * Fsun))

        Dim TauVisDiff As Single = (RefVis ^ 2 - 1) * Math.Exp(-Math.Sqrt(Bioproperties.AlphaVIS) * Kd * Fsun) / _
        ((RefVis * AlphaSoilVis - 1) + RefVis * (RefVis - AlphaSoilVis) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * Kd * Fsun))

        Dim TauNIRDiff As Single = (RefNIR ^ 2 - 1) * Math.Exp(-Math.Sqrt(Bioproperties.AlphaNIR) * Kd * Fsun) / _
       ((RefNIR * AlphaSoilNIR - 1) + RefNIR * (RefNIR - AlphaSoilNIR) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * Kd * Fsun))

        Dim DiffVis As Single = 0.2 : Dim DiffNIR As Single = 0.1 : Dim DirVis As Single = 0.8 : Dim DirNIR As Single = 0.9

        Dim Output As New RnCoefficients_Output
        Dim fvis As Single = 0.48 : Dim fnir = 0.52 ' from Martha's code
        Output.TauSolar = fvis * (DiffVis * TauVisDiff + DirVis * TauVisDir) + fnir * (DiffNIR * TauNIRDiff + DirNIR * TauNIRDir)

        Output.AlbCanopy = fvis * (DiffVis * AlbCanopyVisDiff + DirVis * AlbCanopyVisDir) + fnir * (DiffNIR * AlbCanopyNIRDiff + DirNIR * AlbCanopyNIRDir)
        Output.AlbSoil = fvis * AlphaSoilVis + fnir * AlphaSoilNIR

        Output.TauThermal = Math.Exp(-Bioproperties.AlphaTIR * Fsun) 'from Kustas code

        'Dim meanleaf As single = Fc * 0.55 + (1 - Fc) * 0.3
        Dim TauThermal As Single = Math.Exp(-Math.Sqrt(Bioproperties.AlphaTIR) * Kd * Fsun)
        'Output.TauThermal = 0.99 * (1 - TauThermal)
        Output.TauThermal = TauThermal

        Return Output
    End Function

    Function CoverProperties(ByVal cover As Cover)
        '@Ashish CoverProperties matches the DataGrieView columns indexes 4 to 32 in SETMI Main Window.
        '@Ashish Crop type is seleceted based on its Name from drop-down list in DataGridView and should match cover attribute
        ' from the classified image.
        'classes are based on the NLCD
        ',,,,"α Leaf VIS" "α Leaf NIR" "α Leaf TIR" "α Dead VIS" "α Dead NIR" "α Dead TIR" "fg" "Hc min" "Hc max" "s" "Wc" "ε Soil VIS" "ε Soil NIR" "ε Soil TIR" "Ag" "D" "alphaPT" "Rc ini" "Rc Max"  "Fc VI" "VI Max" "VI Min" "VI Exponent" "Kcbrf VI" "Kcbrf Slope" "Kcbrf Incpt" "Kcbrf Max" "Kcbrf Min" JBB added the last three soil terms here and in all the lists below

        Dim Water = {-999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, -999, "Hardcoded", -999, -999, -999, "Hardcoded", -999, -999, -999, 999}              'Water_bodies
        Dim DefaultCover = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}                 'Annual_crops_associated_with_permanent_crops

        '''' All Bare soils                                                                             'Bare_rocks
        Dim Bare_soil = {0.82, 0.57, 0.95, 0.92, 0.8, 0.95, 1, 0, 0.2, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Sand_and_Gravel = {0.82, 0.57, 0.95, 0.92, 0.8, 0.95, 1, 0, 0.2, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Barren = {0.82, 0.57, 0.95, 0.92, 0.8, 0.95, 1, 0, 0.2, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}

        ''''grass, shrubs, natural grasslnad, pasture, sparcely vegetated areas
        Dim Grass = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Cheat_Grass_and_Other_Weeds = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Desert_Shrubs = {0.82, 0.28, 0.95, 0.42, 0.04, 0.95, 1, 0, 0.6, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Natural_Grassland = {0.82, 0.28, 0.95, 0.42, 0.04, 0.95, 1, 0, 0.6, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}                'Natural Grassland
        Dim Pasture = {0.82, 0.28, 0.95, 0.42, 0.04, 0.95, 1, 0, 0.6, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}                          'Pasture
        Dim Transitional_Woodland_Shrub = {0.85, 0.37, 0.95, 0.72, 0.44, 0.95, 1, 0.6, 0.6, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}    'Transitional_Woodland_Shrub
        Dim Sparsely_vegetated_areas = {0.85, 0.35, 0.95, 0.77, 0.52, 0.95, 1, 0.5, 0.5, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}       'Sparsely_vegetated_areas
        Dim Vegetated_Decadent = {0.85, 0.35, 0.95, 0.77, 0.52, 0.95, 1, 0.5, 0.5, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Upland_Vegetation = {0.85, 0.35, 0.95, 0.77, 0.52, 0.95, 1, 0.5, 0.5, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Upland_Bushes = {0.85, 0.35, 0.95, 0.77, 0.52, 0.95, 1, 0.5, 0.5, 0.02, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}

        ''''All Agricultural surface
        Dim Agriculture = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}                      'Annual_crops_associated_with_permanent_crops
        Dim Alfalfa = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Wheat = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        'Dim Soybean = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 1, 0.05, 0.76, -999, 0.15, 0.25, 0.96, 0.3} 'Modified by JBB
        Dim Soybean = {0.85, 0.2, 0.95, 0.49, 0.13, 0.95, 1, 0, 1, 0.05, 0.76, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999} 'Modified by JBB See Colaizzi et al et al 2012 Agronomy J. 104(2) and Houborg et al. 2009 Ag Forest Met. 149. and Brunsell and Gilles 2002
        'Dim Corn = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 3, 0.05, 0.76, -999, 0.15, 0.25, 0.96, 0.3} 'Modified by JBB
        Dim Corn = {0.85, 0.2, 0.98, 0.49, 0.13, 0.95, 1, 0, 3, 0.05, 0.76, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999} 'Modified by JBB See Colaizzi et al 2012 Agronomy J. 104(2) and Houborg et al. 2009 Ag Forest Met. 149. and Brunsell and Gilles 2002
        Dim Cotton = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0.76, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Dryland_Cotton = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0.76, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Watermelon = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Vineyards = {0.83, 0.35, 0.95, 0.77, 0.52, 0.95, 1, 2.5, 2.5, 0.02, 0.75, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}            'Vineyards
        Dim Sugar_Cane = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 4.0, 0.05, 0.76, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999} 'added by JBB

        'other natural vegetations and forests
        Dim Arrowweed = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 4, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Arundo = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Cottonwood = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Tamarisk = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 4, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Dead_Tamarisk = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 2, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Mesophytes = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Mesquite = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Conifer = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Mixed_Forest = {0.87, 0.49, 0.95, 0.78, 0.53, 0.95, 1, 5, 5, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}               'Mixed_forest
        Dim Coniferous_Forest = {0.89, 0.6, 0.95, 0.84, 0.61, 0.95, 1, 15, 15, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}         'Coniferous_Forest
        Dim Transitional_Forest = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 0.67, 32, 32, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}      'Transitional Forest Savana Amazon with constant height and LAI
        Dim Broad_Leaved_Forest = {0.86, 0.37, 0.95, 0.84, 0.61, 0.95, 1, 10, 10, 0.1, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}       'Broad_Leaved_Forest
        Dim Agro_Forestry_Areas = {0.85, 0.36, 0.95, 0.58, 0.26, 0.95, 1, 1, 2.5, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}      'Agro_Forestry_Areas
        Dim Eastern_Red_Cedar = {0.89, 0.6, 0.95, 0.84, 0.61, 0.95, 1, 5.3, 8.4, 0.2, 8, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999} 'JBB based on conifers from Houborg et al 2009, Heights from Awada et al. 2013, soil and veg Emissivity here and elsewhere from Houborg et al. 2009 and Brunsell and Gilles 2002

        ''''fruit trees
        Dim Pineapple = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}                'Fruit_trees_and_berry_plantations
        Dim Banana = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 4, 4, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Banana_Apple = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 4, 4, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Graviola = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 3, 3, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Sapoti = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 3, 3, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Acerola = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 3, 3, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Sabia = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 3, 3, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Mango = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 3, 3, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Cashew_Giant = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 5, 5, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Cashew_Early = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 5, 5, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Mata = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 5, 5, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Papaya = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 3, 3, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Lemon = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 4, 4, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Tangerine = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 4, 4, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Orange = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 4, 4, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Passion_Fruit = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 2, 2, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Citrus = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 4, 4, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Guava = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 3, 3, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}
        Dim Coconut = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 4, 4, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}

        ' others
        Dim Green_Urban_Areas = {0.87, 0.4, 0.95, 0.84, 0.61, 0.95, 1, 15, 15, 0.1, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}                      'Green_Urban_Areas
        Dim Non_Irrigated_Arable_Land = {0.83, 0.35, 0.95, 0.49, 0.13, 0.95, 1, 0, 0.6, 0.05, 0, -999, 0.15, 0.25, 0.96, 0.3, 1, 1.26, 50, 1000, "Hardcoded", 0.12, 0.68, 1, "Hardcoded", -999, -999, -999, 999}            'Non_Irrigated_Arable_Land

        Dim output
        Select Case cover
            Case Functions.Cover.Water
                output = Water
            Case Functions.Cover.DefaultCover
                output = DefaultCover
            Case Functions.Cover.Bare_Soil
                output = Bare_soil
            Case Functions.Cover.Sand_and_Gravel
                output = Sand_and_Gravel
            Case Functions.Cover.Barren
                output = Barren
            Case Functions.Cover.Cheat_Grass_and_Other_Weeds
                output = Cheat_Grass_and_Other_Weeds
            Case Functions.Cover.Desert_Shrubs
                output = Desert_Shrubs
            Case Functions.Cover.Natural_Grassland
                output = Natural_Grassland
            Case Functions.Cover.Pasture
                output = Pasture
            Case Functions.Cover.Transitional_Woodland_Shrub
                output = Transitional_Woodland_Shrub
            Case Functions.Cover.Sparsely_vegetated_areas
                output = Sparsely_vegetated_areas
            Case Functions.Cover.Vegetated_Decadent
                output = Vegetated_Decadent
            Case Functions.Cover.Upland_Vegetation
                output = Upland_Vegetation
            Case Functions.Cover.Upland_Bushes
                output = Upland_Bushes
            Case Functions.Cover.Agriculture
                output = Agriculture
            Case Functions.Cover.Alfalfa
                output = Alfalfa
            Case Functions.Cover.Wheat
                output = Wheat
            Case Functions.Cover.Soybean
                output = Soybean
            Case Functions.Cover.Corn
                output = Corn
            Case Functions.Cover.Cotton
                output = Cotton
            Case Functions.Cover.Dryland_Cotton
                output = Dryland_Cotton
            Case Functions.Cover.Watermelon
                output = Watermelon
            Case Functions.Cover.Vineyards
                output = Vineyards
            Case Functions.Cover.Arrowweed
                output = Arrowweed
            Case Functions.Cover.Arundo
                output = Arundo
            Case Functions.Cover.Cottonwood
                output = Cottonwood
            Case Functions.Cover.Tamarisk
                output = Tamarisk
            Case Functions.Cover.Dead_Tamarisk
                output = Dead_Tamarisk
            Case Functions.Cover.Mesophytes
                output = Mesophytes
            Case Functions.Cover.Mesquite
                output = Mesquite
            Case Functions.Cover.Conifer
                output = Conifer
            Case Functions.Cover.Mixed_Forest
                output = Mixed_Forest
            Case Functions.Cover.Coniferous_Forest
                output = Coniferous_Forest
            Case Functions.Cover.Transitional_Forest
                output = Transitional_Forest
            Case Functions.Cover.Broad_Leaved_Forest
                output = Broad_Leaved_Forest
            Case Functions.Cover.Agro_Forestry_Areas
                output = Agro_Forestry_Areas
            Case Functions.Cover.Pineapple
                output = Pineapple
            Case Functions.Cover.Banana
                output = Banana
            Case Functions.Cover.Banana_Apple
                output = Banana_Apple
            Case Functions.Cover.Graviola
                output = Graviola
            Case Functions.Cover.Sapoti
                output = Sapoti
            Case Functions.Cover.Acerola
                output = Acerola
            Case Functions.Cover.Sabia
                output = Sabia
            Case Functions.Cover.Mango
                output = Mango
            Case Functions.Cover.Cashew_Giant
                output = Cashew_Giant
            Case Functions.Cover.Cashew_Early
                output = Cashew_Early
            Case Functions.Cover.Mata
                output = Mata
            Case Functions.Cover.Papaya
                output = Papaya
            Case Functions.Cover.Lemon
                output = Lemon
            Case Functions.Cover.Tangerine
                output = Tangerine
            Case Functions.Cover.Orange
                output = Orange
            Case Functions.Cover.Passion_Fruit
                output = Passion_Fruit
            Case Functions.Cover.Citrus
                output = Citrus
            Case Functions.Cover.Guava
                output = Guava
            Case Functions.Cover.Coconut
                output = Coconut
            Case Functions.Cover.Green_Urban_Areas
                output = Green_Urban_Areas
            Case Functions.Cover.Non_Irrigated_Arable_Land
                output = Non_Irrigated_Arable_Land
            Case Functions.Cover.Sugar_Cane 'Added by JBB 
                output = Sugar_Cane 'Added by JBB
            Case Functions.Cover.Eastern_Red_Cedar 'Added by JBB 
                output = Eastern_Red_Cedar 'Added by JBB
            Case Else
                output = DefaultCover
        End Select


        Return output
    End Function
    Function calcRnCoefficients(ByVal Rs As Single, ByVal ClumpSun As Single, ByVal LAI As Single, ByVal Zenith As Single, ByVal Cover As Cover, ByVal Fc As Single, ByRef Bioproperties As Bioproperties, ByVal Patm As Single) As RnCoefficients_Output
        Dim Output As New RnCoefficients_Output
        Dim AlbCanopy As Single
        ' Martha's notes
        'c     *  Computes albedo based on analytical solutions from Goudriaan
        'c     *  1988 and summarized in Campbell & Norman 1998.  These assume
        'c     *  that leaf transmittivity equals leaf reflectivity.  All
        'c     *  we must specify is leaf absorptivity in NIR and VIS
        'c     *  wavebands (ALEAFN, ALEAFV).  Also soil reflectivities
        'c     *  RSOILN and RSOILV.
        'c     *
        'c     *  ZEN-dependent quantities (ALBEDO, TAUBTV and 
        'c     *  TAUBTN) are returned in argument list.

        'Dim AlphaSoilVis As Single = 0.15  ' this was 0.15  from Kustas code , it indicated high soil ref commented out by JBB
        'Dim AlphaSoilNIR As Single = 0.25   ' this was 0.25  from Kustas code , it indicated high soil ref commented out by JBB
        Dim AlphaSoilVis As Single = Bioproperties.EmissSoilVIS 'Added by JBB
        Dim AlphaSoilNIR As Single = Bioproperties.EmissSoilNIR ' Added by JBB

        'Dim DiffVis As single = 0.2 : Dim DiffNIR As single = 0.1 : Dim DirVis As single = 1 - DiffVis : Dim DirNIR As single = 1 - DiffNIR
        'Dim fvis As single = 0.48 : Dim fnir = 0.52 'from Martha's code but another code needs to be checked
        Dim Radcoeff As Radcoefficients = calcRadComps(Rs, Zenith, Patm)
        Dim DiffVis As Single = Radcoeff.DiffVis
        Dim DiffNIR As Single = Radcoeff.DiffNIR
        Dim DirVis As Single = Radcoeff.DirVis
        Dim DirNIR As Single = Radcoeff.DirNIR
        Dim fvis As Single = Radcoeff.fvis
        Dim fnir As Single = Radcoeff.fnir
        Dim fclear As Single = Radcoeff.fclear
        Dim Fsun As Single = ClumpSun * LAI 'See Campbell and Norman 2012 p. 274, This is the fraction of sunlit leaves
        'Dim Xp As single = 1.0
        'Dim Kb As single = Math.Sqrt(Xp ^ 2 + (Math.Tan(Zenith * Math.PI / 180)) ^ 2) / (Xp + 1.774 * (Xp + 1.182) ^ (-0.733)) ' this function is not used see below
        Dim kb As Single = 0
        'Weighted live/dead leaf average properties
        Bioproperties.AlphaVIS = Bioproperties.AlphaLeafVIS * Bioproperties.fg + Bioproperties.AlphaDeadVIS * (1 - Bioproperties.fg) 'Albedos & fg are input in the SETMI GUI, is there a reference for these equations? JBB
        Bioproperties.AlphaNIR = Bioproperties.AlphaLeafNIR * Bioproperties.fg + Bioproperties.AlphaDeadNIR * (1 - Bioproperties.fg) 'Albedos & fg are input in the SETMI GUI, is there a reference for these equations? JBB
        Bioproperties.AlphaTIR = Bioproperties.AlphaLeafTIR * Bioproperties.fg + Bioproperties.AlphaDeadTIR * (1 - Bioproperties.fg) 'Albedos & fg are input in the SETMI GUI, is there a reference for these equations? JBB
        '********************************************************************************************
        '*********Looking through PyTSEB by Hector Nieto and contributors tipped me off that we need to keep absorptivity and emissivity equal
        'Therefore, this code may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
        Eveg = Bioproperties.AlphaTIR 'Added by JBB so that TIR emissivity and absorptivity are equal, this makes sense
        Esoil = Bioproperties.EmissSoilTIR ' Added by JBB so that Emiss Soil TIR is an input
        '************************End PyTSEB influenced code

        'set limits for kd
        Dim Kd As Single = 0
        Kd = -0.0683 * Math.Log(Fsun) + 0.804 'Fit to Fig 15.4 for x=1 'This figure is in Campbell and Norman 1998 JBB, x = 1 would be for a spherical leaf angle distribution a logrithmic regression doesn't fit the data quite right, but it is better than any other option in Excel.

        ''''' Diffuse components '''''

        'Diffuse light canopy reflection coefficients for a deep canopy 
        Dim RefRohVis As Single = (1 - Math.Sqrt(Bioproperties.AlphaVIS)) / (1 + Math.Sqrt(Bioproperties.AlphaVIS))  'Eq. 15.7 ' Equation is in Campbell and Norman 1998 JBB
        Dim RefRohNIR As Single = (1 - Math.Sqrt(Bioproperties.AlphaNIR)) / (1 + Math.Sqrt(Bioproperties.AlphaNIR)) 'Campbell and Norman 1998 Eq. 15.7 JBB
        Dim RefRohLongwave As Single = (1 - Math.Sqrt(Bioproperties.AlphaTIR)) / (1 + Math.Sqrt(Bioproperties.AlphaTIR)) 'Campbell and Norman 1998 Eq. 15.7 JBB
        Dim RefDiffVis As Single = 2 * Kd * RefRohVis / (1 + Kd) 'Eq. 15.8' Equation is in Campbell and Norman 1998 JBB
        Dim RefDiffNIR As Single = 2 * Kd * RefRohNIR / (1 + Kd) 'Campbell and Norman 1998 Eq. 15.8 JBB
        Dim RefDiffLongwave As Single = 2 * Kd * RefRohLongwave / (1 + Kd) 'Campbell and Norman 1998 Eq. 15.8 JBB

        'Diffuse canopy transmission coeff (VIS)
        Dim ExpFac As Single = 0
        ExpFac = (Math.Sqrt(Bioproperties.AlphaVIS) * Kd * Fsun) 'Campbell and Norman 1998 Eq. 15.11 JBB
        Dim TauVisDiff As Single = (RefDiffVis ^ 2 - 1) * Math.Exp(-ExpFac) / _
        ((RefDiffVis * AlphaSoilVis - 1) + RefDiffVis * (RefDiffVis - AlphaSoilVis) * Math.Exp(-2 * ExpFac))  'Eq. 15.11 ' Equation is in Campbell and Norman 1998 JBB

        'Diffuse canopy transmission coeff (NIR)
        ExpFac = Math.Sqrt(Bioproperties.AlphaNIR) * Kd * Fsun 'Campbell and Norman 1998 Eq. 15.11 JBB
        Dim TauNIRDiff As Single = (RefDiffNIR ^ 2 - 1) * Math.Exp(-ExpFac) / _
       ((RefDiffNIR * AlphaSoilNIR - 1) + RefDiffNIR * (RefDiffNIR - AlphaSoilNIR) * Math.Exp(-2 * ExpFac)) 'Eq. 15.11 ' Equation is in Campbell and Norman 1998 JBB

        'Diffuse canopy transmission coeff (longwave)
        'Use deep canopy expression (ignore soil reflection effects)
        ExpFac = Math.Sqrt(Bioproperties.AlphaTIR) * Kd * Fsun 'Campbell and Norman 1998 Eq. 15.6 JBB
        Dim TauThermal As Single = Math.Exp(-ExpFac) 'Eq. 15.6' Equation is in Campbell and Norman 1998, but above we use a Eq. 11 for sparse canopies JBB
        Dim emcpy As Single = 0.99 * (1 - TauThermal) ' based on Bill Eq. 'THIS DOES NOT APPEAR TO BE USED ANYWHERE JBB

        'diffuse radiation surface albedo for a generic crop
        Dim Fact As Single = 0
        Fact = ((RefDiffVis - AlphaSoilVis) / (RefDiffVis * AlphaSoilVis - 1)) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * Kd * Fsun) ' Eq 15.9' Equation is in Campbell and Norman 1998 JBB
        Dim AlbCanopyVisDiff As Single = (RefDiffVis + Fact) / (1 + RefDiffVis * Fact) ' Campbell and Norman 1998 Eq. 15.9 JBB

        Fact = (RefDiffNIR - AlphaSoilNIR) / (RefDiffNIR * AlphaSoilNIR - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * Kd * Fsun) ' Eq 15.9' Equation is in Campbell and Norman 1998 JBB
        Dim AlbCanopyNIRDiff As Single = (RefDiffNIR + Fact) / (1 + RefDiffNIR * Fact) ' Campbell and Norman 1998 Eq. 15.9 JBB

        Dim TauVisDir As Single
        Dim TauNIRDir As Single
        'check if we low zenith angle. it is indicated by Martha that Diffuse is all we can compute in this case
        '        If Math.Cos(ToRadians(Zenith)) <= 0.01 Then 'This is a zenith angle of 89.43 degrees or greater JBB
        If Math.Cos(ToRadians(Zenith)) <= 0.1 Then 'Modified by JBB to match the logic in RadComps

            AlbCanopy = fvis * (DiffVis * AlbCanopyVisDiff) + fnir * (DiffNIR * AlbCanopyNIRDiff) 'Colaizzi et al 2012 Eq. 11, but neglecting direct JBB NOTE AT LOW ZENITH FVIS = FNIR = 0.5
            TauVisDir = 0
            TauNIRDir = 0
            Dim AlbCanopyVisDir As Single = 0
        Else

            ' Beam Components
            'Direct beam extinction coeff (spher. LAD)
            kb = 0.5 / Math.Cos(ToRadians(Zenith)) 'Campbell and Norman 1998 Eq. 15.3

            ' Direct beam canopy reflection coefficients for a deep canopy
            RefRohVis = (1 - Math.Sqrt(Bioproperties.AlphaVIS)) / (1 + Math.Sqrt(Bioproperties.AlphaVIS))  'Eq. 15.7 'This Eq is in Campbell and Norman 1998 JBB, this is actually redundent from above
            RefRohNIR = (1 - Math.Sqrt(Bioproperties.AlphaNIR)) / (1 + Math.Sqrt(Bioproperties.AlphaNIR)) 'Campbell and Norman 1998 Eq. 15.7 JBB, this is actually redundent from above
            Dim RefDirVis As Single = 2 * kb * RefRohVis / (1 + kb) 'Eq. 15.8 'This Eq is in Campbell and Norman 1998 JBB
            Dim RefDirNIR As Single = 2 * kb * RefRohNIR / (1 + kb) 'Campbell and Norman 1998 Eq. 15.8 JBB

            'Direct beam radiation surface albedo for a generic canopy
            Fact = (RefDirVis - AlphaSoilVis) / (RefDirVis * AlphaSoilVis - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaVIS) * kb * Fsun) ' Campbell and Norman 1998 Eq. 15.9
            Dim AlbCanopyVisDir As Single = (RefDirVis + Fact) / (1 + RefDirVis * Fact) ' Campbell and Norman 1998 Eq. 15.9

            Fact = (RefDirNIR - AlphaSoilNIR) / (RefDirNIR * AlphaSoilNIR - 1) * Math.Exp(-2 * Math.Sqrt(Bioproperties.AlphaNIR) * kb * Fsun) ' Campbell and Norman 1998 Eq. 15.9
            Dim AlbCanopyNIRDir As Single = (RefDirNIR + Fact) / (1 + RefDirNIR * Fact) ' Campbell and Norman 1998 Eq. 15.9

            'weighted average albedo
            AlbCanopy = fvis * (DiffVis * AlbCanopyVisDiff + DirVis * AlbCanopyVisDir) + fnir * (DiffNIR * AlbCanopyNIRDiff + DirNIR * AlbCanopyNIRDir) 'Colaizzi et al 2012 Eq. 11 JBB

            'Direct beam+scattered canopy transmission coeff (visible) 
            ExpFac = Math.Sqrt(Bioproperties.AlphaVIS) * kb * Fsun 'Campbell and Norman 1998 Eq. 15.11,  JBB
            TauVisDir = ((RefDirVis ^ 2 - 1) * Math.Exp(-ExpFac)) / ((RefDirVis * AlphaSoilVis - 1) + RefDirVis * (RefDirVis - AlphaSoilVis) * Math.Exp(-2 * ExpFac)) 'Eq. 15.11 ' Equation is in Campbell and Norman 1998 JBB

            'Direct beam+scattered canopy transmission coeff (NIR) 	
            ExpFac = Math.Sqrt(Bioproperties.AlphaNIR) * kb * Fsun 'Campbell and Norman 1998 Eq. 15.11,  JBB
            TauNIRDir = ((RefDirNIR ^ 2 - 1) * Math.Exp(-ExpFac)) / ((RefDirNIR * AlphaSoilNIR - 1) + RefDirNIR * (RefDirNIR - AlphaSoilNIR) * Math.Exp(-2 * ExpFac)) 'Eq. 15.11 ' Equation is in Campbell and Norman 1998 JBB

        End If

        Output.TauSolar = fvis * (DiffVis * TauVisDiff + DirVis * TauVisDir) + fnir * (DiffNIR * TauNIRDiff + DirNIR * TauNIRDir) 'Colaizzi et al 2012 with fsc=1 following p220 1st col 2nd to last paragraph JBB
        Output.TauThermal = TauThermal
        Output.AlbSoil = fvis * AlphaSoilVis + fnir * AlphaSoilNIR 'Colaizzi et al 2012 EQ 14 Just weighting JBB
        Output.AlbCanopy = AlbCanopy
        Output.fclear = fclear

        Return Output
    End Function

    Function calcRadComps(ByVal Rs As Single, ByVal Zenith As Single, ByVal Patm As Single) As Radcoefficients
        '*****These are based on Weiss and Norman 1985 ****** JBB
        'The code has been modified to follow the original formulation with some modifications to better fit current
        'solar constants kinda like PyTSEB did, therefore some parts of this (as noted) may fall under the PyTSEB General Public License see:https://github.com/hectornieto/pyTSEB

        Dim ration As Single = 1
        Dim fclear As Single
        Dim fvis As Single
        Dim fnir As Single
        Dim diffnir As Single
        Dim diffvis As Single
        Dim dirvis As Single
        Dim dirnir As Single
        Dim P0 As Single = 1013.25 'Atmospheric pressure at sea level follwing W&N85, but converting to mbar rather than kPa

        If Math.Cos(ToRadians(Zenith)) < 0.1 Then
            fclear = ration
            fvis = 0.5
            fnir = 0.5
            diffvis = 1
            diffnir = 1
            dirvis = 0
            dirnir = 0
        Else
            'Calculate potential (clear-sky) visible and NIR solar components
            'Correct for curvature of atmos in airmas
            'Dim airmas As Single = (Math.Sqrt((Math.Cos(ToRadians(Zenith))) ^ 2 + 0.0025) - Math.Cos(ToRadians(Zenith))) / 0.00125'Commented out by JBB

            '//Correct for refraction(good to 89.5 deg.)
            'airmas = airmas - 2.8 / (90 - Zenith) ^ 2'Commented out by JBB

            Dim airmas As Single = 1 / Math.Cos(ToRadians(Zenith)) 'W&N85 Eq 2, PyTSEB also does this

            'Dim potbm1 As Single = 600 * Math.Exp(-0.16 * airmas)'Commented out by JBB
            Dim potbm1 As Single = 600 * Math.Exp(-0.185 * Patm / P0 * airmas)  'Eq 1., without Cos Zenith, yet following SETMI
            'Dim potvis As Single = (potbm1 + (600 - potbm1) * 0.4) * Math.Cos(ToRadians(Zenith)) 'An unnecessary addition of potmb1 and potdif
            'Dim potdif As Single = (600 - potbm1) * 0.4 * Math.Cos(ToRadians(Zenith)) 'Commented out by JBB
            Dim potdif1 As Single = (600 - potbm1) * 0.4 * Math.Cos(ToRadians(Zenith)) 'Eq 3, no changes, just renamed potdif to potdif1 folowing SETMI whihc originally didn't have the cos(zenith) in potbm1 before this step, the original paper does, however, but including it first seems to double count it
            potbm1 = potbm1 * Math.Cos(ToRadians(Zenith)) 'Eq 1, adding Cos(Zenith) after PotDif is calc'd add by JBB
            Dim potvis = potbm1 + potdif1 'added by JBB for simpilicity
            Dim u As Single = 1.0 / Math.Cos(ToRadians(Zenith)) 'Same as airmas
            Dim axlog As Single = Math.Log10(u) 'Eq 6
            Dim a As Single = (10 ^ (-1.195 + 0.4459 * axlog - 0.0345 * axlog ^ 2)) 'Eq 6
            Dim watabs As Single = 1320 * a 'Eq 6
            'Dim potbm2 As Single = 720 * Math.Exp(-0.05 * airmas) - watabs'Commented out by JBB
            Dim potbm2 As Single = (720 * Math.Exp(-0.06 * Patm / P0 * airmas) - watabs)  'Eq. 4 added by JBB, adding Cos(Zenith) later like SETMI did originally

            If (potbm2 < 0) Then potbm2 = 0

            'Dim eval As Single = (720 - potbm2 - watabs) * 0.54 * Math.Cos(ToRadians(Zenith))'commented out by JBB 
            Dim potdif2 As Single = (720 - potbm2 - watabs) * 0.6 * Math.Cos(ToRadians(Zenith)) 'Eq 5
            potbm2 = potbm2 * Math.Cos(ToRadians(Zenith)) 'Eq 4 adding Cos(zenith) after PotDif is calc'd add by JBB
            'Dim potnir As Single = eval + potbm2 * Math.Cos(ToRadians(Zenith))'commented out by JBB
            Dim potnir As Single = potdif2 + potbm2 'added by JBB for simplicity
            fclear = Rs / (potvis + potnir)
            fclear = Math.Min(1.0, fclear)

            ' //Partition SDN into VIS and NIR
            fvis = potvis / (potvis + potnir)
            fnir = potnir / (potvis + potnir)

            '//Estimate direct beam and diffuse fractions in VIS and NIR wavebands
            'Dim fb1 As Single = potbm1 * Math.Cos(ToRadians(Zenith)) / potvis
            Dim fb1 As Single = potbm1 / potvis 'Since CosZenith is added in above now JBB
            'Dim fb2 As Single = potbm2 * Math.Cos(ToRadians(Zenith)) / potnir
            Dim fb2 As Single = potbm2 / potnir 'Since CosZenith is added in above now JBB

            Dim ratiox As Single = fclear

            If (fclear > 0.9) Then ratiox = 0.9
            dirvis = fb1 * (1 - ((0.9 - ratiox) / 0.7) ^ 0.6667)

            If (fclear > 0.88) Then ratiox = 0.88
            dirnir = fb2 * (1 - ((0.88 - ratiox) / 0.68) ^ 0.6667)

            dirvis = Math.Max(0.0, dirvis)
            dirnir = Math.Max(0.0, dirnir)
            dirvis = Math.Min(fb1, dirvis)
            dirnir = Math.Min(fb2, dirnir)

            If (dirvis < 0.01 And dirnir > 0.01) Then dirvis = 0.011
            If (dirnir < 0.01 And dirvis > 0.01) Then dirnir = 0.011

            diffvis = 1.0 - dirvis
            diffnir = 1.0 - dirnir
        End If

        Dim output As New Radcoefficients
        output.fnir = fnir
        output.fvis = fvis
        output.DiffNIR = diffnir
        output.DiffVis = diffvis
        output.DirNIR = dirnir
        output.DirVis = dirvis
        output.fclear = fclear

        Return output

    End Function

    ' Function calcTcomponents(ByVal Temp As Single, ByRef Tcanopy As Single, ByRef Tsoil As Single, ByVal Tair As Single, ByVal Resistances As Resistances_Output, ByVal fTheta As Single, ByVal W As W_Output, ByVal EnergyComponents As EnergyComponents_Output, ByVal Cover As Cover) As TInitial_Output
    Function calcTcomponents(ByVal Temp As Single, ByRef Tcanopy As Single, ByRef Tsoil As Single, ByVal Tair As Single, ByVal Resistances As Resistances_Output, ByVal fTheta As Single, ByVal W As W_Output, ByVal EnergyComponents As EnergyComponents_Output, ByVal Cover As Cover, ByVal LAI As Single, ByVal FractionCover As Single) As TInitial_Output
        'Tair -= 273.15
        'Temp -= 273.15

        'multiply all equations by Rsoil to avoid division by zero for bare soil condition (from Matha's code)
        Dim Numer As Single : Dim Deno As Single ' These calculations are equivalent to Eq. A.7 in Norman et al 1995 JBB
        'Numer = (Tair / Resistances.Ra + Temp / (Resistances.Rsoil * (1 - fTheta)) + Resistances.Rx / W.Cp_TSM * EnergyComponents.Hcanopy * (1 / Resistances.Ra + 1 / Resistances.Rsoil + 1 / Resistances.Rx))
        'Deno = (1 / Resistances.Ra + 1 / Resistances.Rsoil + fTheta / (Resistances.Rsoil * (1 - fTheta)))
        Numer = (Tair * Resistances.Rsoil / Resistances.Ra + Temp / (1 - fTheta) + EnergyComponents.Hcanopy * Resistances.Rx / W.CpRho * (Resistances.Rsoil / Resistances.Ra + 1 + Resistances.Rsoil / Resistances.Rx))
        'xnum=  (tak*Entrada->rs/Entrada->ra)+(tradk/((1.-ftheta)))+ (Entrada->hc*(0.5*Entrada->rx)/rhocp)* ((Entrada->rs/Entrada->ra)+(Entrada->rs/(0.5*Entrada->rx))+(1.0));
        Deno = (Resistances.Rsoil / Resistances.Ra + 1 + fTheta / (1 - fTheta))
        Dim TcLin As Single = Numer / Deno ' These calculations are equivalent to Eq. A.7 in Norman et al 1995 JBB
        'TcLin = TcLin + 273.15

        Dim Td As Single = TcLin * (1 + Resistances.Rsoil / Resistances.Ra) - Resistances.Rx / W.CpRho * EnergyComponents.Hcanopy * (1 + Resistances.Rsoil / (Resistances.Rx) + Resistances.Rsoil / Resistances.Ra) - Tair * Resistances.Rsoil / Resistances.Ra 'Norman et al 1995 Eq. A.12 JBB

        Dim DeltaTc As Single = (Temp ^ 4 - fTheta * TcLin ^ 4 - (1 - fTheta) * Td ^ 4) / ((1 - fTheta) * 4 * Td ^ 3 * (1 + Resistances.Rsoil / Resistances.Ra) + fTheta * 4 * TcLin ^ 3) 'Norman et al 1995 Eq A.11 w/ n = 4 (see p 265 N95 descript. of EQ 1). JBB
        Tcanopy = TcLin + DeltaTc 'Norman et al 1995 Eq. A.18 JBB

        If LAI <= LAIMin Or FractionCover <= FcoverMin Then 'added by JBB
            Tcanopy = Tair 'Following an email from Bill Kustas 7/6/2016
        End If

        Tsoil = calcTsoilFromTcanopy(Temp, fTheta, Tcanopy)
        Dim Tac As Single
        'Tac = Tcanopy - EnergyComponents.Hcanopy * Resistances.Rx / W.Cp_TSM ''''' based on Kustas code'This was a rearrangement of Norman et al 1995 Eq. A.2. JBB
        'Tac = (Temp * Resistances.Rsoil / Resistances.Ra + Tsoil + Tcanopy * Resistances.Rsoil / Resistances.Rx) / (Resistances.Rsoil / Resistances.Ra + 1 + Resistances.Rsoil / Resistances.Rx) ''''based on EQ. A4 Norman et al. 1995 '<-- Trad has been substituted for Tair, can we assume Trad = Tair? JBB commented out by JBB
        Tac = (Tair * Resistances.Rsoil / Resistances.Ra + Tsoil + Tcanopy * Resistances.Rsoil / Resistances.Rx) / (Resistances.Rsoil / Resistances.Ra + 1 + Resistances.Rsoil / Resistances.Rx) ''''based on EQ. A4 Norman et al. 1995 '<-- subsituted Tair in for Temp so the eq. is correct by JBB added by JBB
        Dim Output As New TInitial_Output ' all in K
        Output.Tac = Tac
        Output.Tcanopy = Tcanopy
        Output.Tsoil = Tsoil
        'If LAI <= LAIMin Or FractionCover <= FcoverMin Then
        '    Output.Tsoil = Temp 'overwrites output for Tsoil, the other temps are left alone and just not used later
        'End If
        Return Output
        If Resistances.Rsoil = 0.0 Then
            Tac = Tac
        End If
    End Function

    Function calcTsoilFromTcanopy(ByVal Temp As Single, ByVal fTheta As Single, ByVal Tcanopy As Single) As Single
        'trp is in oK and tcano is in oC
        'Temp += 273.15
        'Tcanopy += 273.15
        Dim Tsoil As Single
        If Temp ^ 4 - fTheta * Tcanopy ^ 4 < 0 Then ' This appears to be a convenience to avoid a fourth root of a negative number, but if the differences are large this simplification diverges, a better option would be to use an absolute value!, either way in this range Tc becomes a very cold number (below freezing) really if ftheta is low Ts should be really close to Trad! JBB
            'Tsoil = (Temp - 273.15 - fTheta * (Tcanopy - 273.15)) / (1 - fTheta) 'Is the -273.15 erroneous ??? JBB YES!!!
            Tsoil = (Temp - fTheta * (Tcanopy)) / (1 - fTheta)
        Else
            Tsoil = ((Temp ^ 4 - fTheta * Tcanopy ^ 4) / (1 - fTheta)) ^ (1 / 4) 'This is a rearranged form of Norman et al 1995 Eq 1 w/ n = 4 JBB
        End If
        Return Tsoil  'in K
    End Function

    Function calcTcanopyFromTsoil(ByVal Temp As Single, ByVal fTheta As Single, ByVal Tsoil As Single) As Single
        'Tsoil += +273.15
        Dim Tcanopy As Single
        If Tsoil - Temp > 15 Then 'This logic appears to be accomplishing a similar goal as the following logic, i.e. preventing a fourth root of a negative number JBB if Trad = 0C, and Tsoil > 15 C and ftheta is low this could happen.
            Tsoil = Temp + 15
        End If

        If Temp ^ 4 - (1 - fTheta) * Tsoil ^ 4 < 0 Then
            Tcanopy = (Temp - (1 - fTheta) * Tsoil) / fTheta ' This appears to be a convenience to avoid a fourth root of a negative number, but if the differences are large this simplification diverges, a better option would be to use an absolute value!, either way in this range Tc becomes a very cold number (below freezing) really if ftheta is low Ts should be really close to Trad! JBB
        Else
            Tcanopy = ((Temp ^ 4 - (1 - fTheta) * Tsoil ^ 4) / fTheta) ^ (1 / 4) 'This is a rearrangement of Eq A.1 in Kustas and Norman 1999 JBB
        End If

        Return Tcanopy  'in oK
    End Function

    Function calcW(ByVal Tair As Single, ByVal Ea As Single, ByVal P As Single, ByVal Es As Single, ByVal Tr As Single) As W_Output
        'Dim Rho As Single = 1 / (287.04 * Tair) * (P - 0.3783 * Ea) * 100 'this is to Eq 3.4 + 3.5 with a unit conversion from mbar to pa in BrutseartJBB
        Dim Rho As Single = 1 / (287.04 * Tair) * (P - 0.378 * Ea) * 100 'this is to Eq 3.4 + 3.5 with a unit conversion from mbar to pa in BrutseartJBB JBB Changed constant to 0.378 to match Brutseart
        'Dim Lambda As Single = 2501300 - 2366 * (Tair - 273.15) 'WHERE IS THIS EQ FROM? JBB
        Dim Lambda As Single = 2501000 - 2361 * (Tair - 273.15) 'Added by JBB from Ham 2005 Eq. 10
        Dim Rhov As Single = 0.622 * Ea * 100 / (287.04 * Tair) 'this is similar to Eq 3.5 of Brutseart with a unit conversion from mbar to pa JBB
        Dim Q As Single = Rhov / Rho 'Brutseart 1982 Eq 3.2 JBB
        Dim Cp0 As Single = 1005 'Cp of dry air see Table 3.1 in Brutseart 1982 JBB
        Dim Output As New W_Output
        Output.Cp2 = Cp0 * (1 + 0.84 * Q) 'specific heat of moist air Equation 3.31 Brutseart 1982 JBB
        Output.Gamma = Output.Cp2 * P / (0.622 * Lambda) 'Brutseart 1982 Eq. 10.10 'JBB
        Dim Delta As Single = 373.15 * Es * (13.3185 - 3.952 * Tr - 1.9335 * Tr ^ 2 - 0.5196 * Tr ^ 3) / (Tair ^ 2) 'Eq 3.24b from Brutseart 1982 JBB
        Output.W = Delta / (Output.Gamma + Delta) 'The energy partition from PT JBB
        Output.CpRho = Output.Cp2 * Rho
        Output.Cp2P = Output.Cp2 * P
        Output.Lambda = Lambda
        Output.Rho = Rho
        Return Output
    End Function

    Function calcW2(ByVal Tcanopy As Single, ByVal Cp2P As Single, ByVal Lambda As Single, ByVal specialstablecase As Integer) As Single

        Dim Gamma As Single
        If specialstablecase = 1 Then ' In TSEB this is always hard coded to be 1 JBB
            'Lambda = 2501300 - 2366 * (Tcanopy) 'WHERE IS THIS RELATIONSHIP FROM? JBB - IT SEEMS TO BE COMMON
            Lambda = 2501000 - 2361 * (Tcanopy - 273.15) 'Added by JBB from Ham 2005 Eq. 10
        End If
        Gamma = Cp2P / (0.622 * Lambda) 'Brutseart 1982 Eq. 10.10 'JBB

        Dim Tres As Single = 1 - 373.15 / (Tcanopy) 'Brutseart P. 42 JBB
        Dim Esp As Single = 1013.25 * Math.Exp(13.3185 * Tres - 1.976 * Tres ^ 2 - 0.6445 * Tres ^ 3 - 0.1299 * Tres ^ 4) 'Eq 3.24a from Brutseart 1982 JBB
        Dim Delta As Single = 373.15 * Esp * (13.3185 - 3.952 * Tres - 1.9335 * Tres ^ 2 - 0.5196 * Tres ^ 3) / (Tcanopy) ^ 2 'Eq 3.24b from Brutseart 1982 JBB

        Return Delta / (Delta + Gamma) 'The energy partition from PT JBB
    End Function

    Function calcDelta2(ByVal Tcanopy As Single, ByVal Cp2P As Single, ByVal Lambda As Single) As Single
        'Calculates Esat/Temp Gradient for PM 'Based on CalcW2 added by JBB
        Dim Tres As Single = 1 - 373.15 / (Tcanopy) 'Brutseart P. 42 JBB
        Dim Esp As Single = 1013.25 * Math.Exp(13.3185 * Tres - 1.976 * Tres ^ 2 - 0.6445 * Tres ^ 3 - 0.1299 * Tres ^ 4) 'Eq 3.24a from Brutseart 1982 JBB
        Dim Delta As Single = 373.15 * Esp * (13.3185 - 3.952 * Tres - 1.9335 * Tres ^ 2 - 0.5196 * Tres ^ 3) / (Tcanopy) ^ 2 'Eq 3.24b from Brutseart 1982 JBB
        Return Delta
    End Function
    Function calcGammaStar2(ByVal Tcanopy As Single, ByVal Cp2P As Single, ByVal Lambda As Single, ByVal Rcan As Single, ByVal Ra As Single) As Single
        'Calculates Gamma Star for PM 'Based on CalcW2 added by JBB
        Dim Gamma As Single
        Dim GammaStar As Single
        Lambda = 2501000 - 2361 * (Tcanopy - 273.15) 'Added by JBB from Ham 2005 Eq. 10
        Gamma = Cp2P / (0.622 * Lambda) 'Brutseart 1982 Eq. 10.10 'JBB
        GammaStar = Gamma * (1 + Rcan / Ra) 'P 480 of Colaizzi et al 2014
        Return GammaStar
    End Function

    Function calcU2(ByVal U As Single, ByVal Z As Single, ByVal windspeedindex As Integer) As Single
        If windspeedindex > -1 Then
            Return 4.87 * U / Math.Log(67.8 * Z - 5.42) 'ASCE STD MANUAL EQ 33 (ASSUMES GROUND COVER IS 10 CM TALL) JBB
        Else
            Return U
        End If
    End Function

    Function calcSunZenith(ByVal RecordDate As DateTime, ByVal SiteLatitude As Single, ByVal SiteLongitude As Single, ByVal StandardLongitude As Single) As Single
        SiteLatitude = ToRadians(SiteLatitude)
        ' Dim TimeHour As single = 11.0 'Commented out by JBB, Time was fixed for 11:00 am JBB Can't we make this an input?
        Dim TimeHour As Single = RecordDate.Hour 'Added by JBB so that time can be grabbed from image file name JBB
        'Dim TimeMinute As single = 0 'Commented out by JBB, Time was fixed for 11:00 am JBB Can't we make this an input?
        Dim TimeMinute As Single = RecordDate.Minute 'Added by JBB so that the time can be obtained from the image file name JBB
        Dim B As Single = 2 * Math.PI * (RecordDate.DayOfYear - 81.25) / 365 'This is a radians version of Eq. 1.4.2 in Duffie and Beckman 2013 Solar Engineering of Thermal Processes JBB
        Dim E As Single = 9.87 * Math.Sin(2 * B) - 7.53 * Math.Cos(B) - 1.5 * Math.Sin(B) 'This Eq is in CMU Neale's RS Class Notes from Spring 2009. JBB
        Dim Delta As Single = ToRadians(23.45 * Math.Sin(2 * Math.PI * (284 + RecordDate.DayOfYear) / 365)) 'Solar Declination Eq. 1.6.1A in Duffie and Beckman 2013 Solar Engineering of Thermal Processes JBB
        'Dim SolarTime As single = (RecordDate.Hour + RecordDate.Minute / 60) + 4 * (StandardLongitude - SiteLongitude) / 60 + E / 60
        Dim SolarTime As Single = (TimeHour + TimeMinute / 60) + 4 * (StandardLongitude - SiteLongitude) / 60 + E / 60 'This is a decimal hour version of Eq 1.5.2 in Duffie and Beckman 2013 Solar Engineering of Thermal Processes JBB
        Dim Omega As Single = ToRadians((SolarTime - 12) * 15) 'See p. 13 of Duffie and Beckman 2013 or CMU Neale's RS Class Notes from Spring 2009. JBB
        Dim cosZ As Single = Math.Sin(Delta) * Math.Sin(SiteLatitude) + Math.Cos(Delta) * Math.Cos(SiteLatitude) * Math.Cos(Omega) 'Eq 1.6.5 from Duffie and Beckman 2013 Solar Engineering of Thermal Processes JBB
        calcSunZenith = ToDegrees(Math.Acos(cosZ)) ' Radians to Degrees JBB
    End Function

    'Function calcTInitial(ByVal Temp As Single, ByVal Tair As Single, ByVal Ftheta As Single) As TInitial_Output
    Function calcTInitial(ByVal Temp As Single, ByVal Tair As Single, ByVal Ftheta As Single, ByVal LAI As Single, ByVal Fc As Single) As TInitial_Output 'added by JBB for bare soil
        Dim Tcanopy As Single = (Tair + Temp) / 2 'in oK 'Why is this okay? JBB
        Dim Tsoil As Single = 0
        'tsoil is in oC

        If LAI <= LAIMin Or Fc < FcoverMin Then 'added by JBB for bare soil
            Tcanopy = Tair 'Following an email from Bill Kustas 7/6/2016
        End If


        If (Temp ^ 4 - Ftheta * Tcanopy ^ 4) < 0 Then
            Tsoil = (Temp - Ftheta * Tcanopy) / (1 - Ftheta) ' This appears to be a convenience to avoid a fourth root of a negative number, but if the differences are large this simplification diverges, a better option would be to use an absolute value!, either way in this range Tc becomes a very cold number (below freezing) really if ftheta is low Ts should be really close to Trad! JBB
        Else
            Tsoil = (((Temp ^ 4 - Ftheta * Tcanopy ^ 4) / (1 - Ftheta)) ^ (1 / 4)) ' This is the more correct version JBB
        End If
        'If LAI <= LAIMin Or Fc < FcoverMin Then 'added by JBB for bare soil
        '    Tsoil = Temp
        'End If
        'check the computed tsoil
        If Tsoil - Temp > 15 Then 'Limits Tsoil to not be too much bigger than Trad
            Tsoil = Temp + 15
            If (Temp ^ 4 - (1 - Ftheta) * (Tsoil) ^ 4) < 0 Then
                Tcanopy = (Temp - (1 - Ftheta) * Tsoil) / Ftheta ' This appears to be a convenience to avoid a fourth root of a negative number, but if the differences are large this simplification diverges, a better option would be to use an absolute value!, either way in this range Tc becomes a very cold number (below freezing) really if ftheta is low Ts should be really close to Trad! JBB
            Else
                Tcanopy = ((Temp ^ 4 - (1 - Ftheta) * Tsoil ^ 4) / Ftheta) ^ (1 / 4) 'This is a rearrangement of Eq A.1 in Kustas and Norman 1999 JBB
            End If
        End If

        Dim TInitial As New TInitial_Output
        TInitial.Tcanopy = Tcanopy        'in K
        TInitial.Tsoil = Tsoil          'in oK
        TInitial.Tac = 0

        Return TInitial
    End Function

    Function calcAg(ByVal NDVI As Single, ByRef Bioproperties As Bioproperties) As Single 'JBB Changed the reference for this function to call the new fucntion CalcAg2 JBB
        Dim Ag As Single = 0.15 '0.3 'THIS 0.15 ISN'T EVEN USED!!!!!!!!!! JBB IT IS OVER WRITTEN LATER IN THIS FUNCTION

        'timeratio = Math.Abs(timeop - 12.5) / 12.5 'This is Kustas and Norman 1999 (see Kustas et al 1998) Eq A.11 with hard programed solar noon = 12.5. JBB

        Dim TimeRatio As Single = 0.2 ' WHY IS THIS TIME RATIO HARD PROGRAMMED AT 0.2? JBB
        If TimeRatio < 0.27 Then 'Kustas and Norman 1999 say <0.3 P 26 top of second col. JBB
            'Ag = 0.3 'Brutseart 1982 Eq 6.42 cites an average of 0.3, Norman et al 1995 and following papers used 0.35 near solar noon! JBB
            'ag = 0.583 * Math.Exp(-2.13 * NDVI(countR, countC))
            Ag = Bioproperties.SoilHtFlxAg 'Added by JBB
        Else
            'Ag = 0.925 - 2.184 * TimeRatio 'Following Kustas et al 1998 p 124 (or Kustas and Norman 1999 p 26), but what data was used for this regression? JBB
            Ag = Bioproperties.SoilHtFlxAg 'Added by JBB
        End If

        Return Ag
    End Function
    Function calcAg2(ByVal LongitudeDeg As Single, ByVal TimeZone As Single, ByVal DoY As Integer, ByVal ImageDecimalHour As Single, ByRef Bioproperties As Bioproperties) As Single 'Added by JBB to replace CalcAg, Includes Solar Noon Calc
        Dim Ag As Single
        Dim TimeRatio As Single
        Dim SolarNoonHours As Single
        SolarNoonHours = CalcSolarNoon(LongitudeDeg, TimeZone, DoY) * 24
        TimeRatio = Math.Abs(ImageDecimalHour - SolarNoonHours) / SolarNoonHours 'This is Kustas and Norman 1999 (see Kustas et al 1998) Eq A.11 JBB

        If TimeRatio < 0.3 Then 'Kustas and Norman 1999 say <0.3 P 26 top of second col. JBB
            'Ag = 0.3 'Brutseart 1982 Eq 6.42 cites an average of 0.3, Norman et al 1995 and following papers used 0.35 near solar noon! JBB
            'Ag = 0.35 'Added by JBB to follow TSEB literature see note in line above
            Ag = Bioproperties.SoilHtFlxAg 'Added by JBB
        Else
            'Ag = 0.925 - 2.184 * TimeRatio 'Following Kustas et al 1998 p 124 (or Kustas and Norman 1999 p 26), but what data was used for this regression?, it matches the 0.3 well when TimeRatio is just > 0.3 JBB
            Ag = Bioproperties.SoilHtFlxAg 'Added by JBB
        End If

        Return Ag
    End Function
    Function CalcSolarNoon(ByVal LongitudeDeg As Single, ByVal TimeZone As Integer, ByVal DoY As Integer) As Single 'by JBB Calculates Solar Noon in fractional days Based on the ASCE Std. ETr Manual JBB
        'Both longDeg and Time zone are negative to the west, UPDATED 8/11/2016 to match Solar Zenith Angle Calculation
        'Dim b As Single = 2 * Math.PI * (DoY - 81) / 364 'Eq. 58
        Dim B As Single = 2 * Math.PI * (DoY - 81.25) / 365 'added by JBB for consistency This is a radians version of Eq. 1.4.2 in Duffie and Beckman 2013 Solar Engineering of Thermal Processes JBB
        'Dim Sc As Single = 0.1645 * Math.Sin(2 * b) - 0.1255 * Math.Cos(b) - 0.025 * Math.Sin(b) 'Eq. 57
        Dim E As Single = 9.87 * Math.Sin(2 * B) - 7.53 * Math.Cos(B) - 1.5 * Math.Sin(B) 'This Eq is in CMU Neale's RS Class Notes from Spring 2009. JBB
        Dim Lonz As Single = TimeZone * 15
        'Dim SolarNoonHour As Single = 12 - Sc + (Lonz - LongitudeDeg) / 15 'Rearrangement of Eq 55, with longitudes positive east of 0
        Dim SolarNoonHour As Single = 12 - E / 60 + (Lonz - LongitudeDeg) / 15 'Ham 2005 Eq. 42 , with longitudes positive east of 0
        Dim SolarNoon As Single = SolarNoonHour / 24 'Solar noon in decimal days
        CalcSolarNoon = SolarNoon
    End Function

    Function calcKcbReflectance(ByVal Cover As Cover, ByVal SAVI As Single, ByVal ETReferenceType As ETReferenceType, ByVal NDVI As Single, ByVal KcbVI As KcbVI, ByVal InputSlope As Single, InputIntercept As Single, InputKcbMax As Single, InputKcbMin As Single)
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '!!!!!!  UPDATES IN THIS FUNCTION MUST BE ACCOMPANIED BY UPDATES IN FUNCTION: FindKcbVI() !!!!!!!
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Dim Kcint As Single = 0 'added by JBB
        If KcbVI = KcbVI.Hardcoded Then
            Select Case ETReferenceType
                Case Functions.ETReferenceType.Short_Grass
                    Select Case Cover
                        Case Functions.Cover.Corn
                            'Return SAVI * 1.835 - 0.034 'Matches Hatim's dissertation JBB
                            Return 1.97 * SAVI - 0.11 'From Email from Ivo Zution 9/1/2017 - likely computed by Isidro Campos, should be equivalent to the tall reference version of Campos et al 2017
                        'Return 0
                        Case Functions.Cover.Soybean
                            'Return SAVI * 1.638 + 0.003 'Matches Hatim's dissertation JBB
                            Return 0
                        Case Functions.Cover.Cotton
                            Return SAVI * 1.587 + 0.007
                        Case Functions.Cover.Dryland_Cotton
                            Return SAVI * 1.587 + 0.007
                        Case Functions.Cover.Sugar_Cane '@Ashish: Added Sugarcane Kcb calculation capablity
                            Return SAVI * 1.5567 - 0.0356
                        Case Else
                            Return 0
                    End Select
                Case Functions.ETReferenceType.Tall_Alfalfa 'Needs to be populated
                    Select Case Cover
                        Case Functions.Cover.Corn
                            'Kcint = SAVI * 1.416 + 0.0174 'Following Bausch added by JBB
                            'Kcint = SAVI * 1.414 - 0.02 'Isidro's new numbers
                            Return SAVI * 1.414 - 0.02 'Isidro's new numbers
                        'Return Limit(Kcint, 0.12, 0.95) 'added by JBB

                        Case Functions.Cover.Soybean
                            'Kcint = SAVI * 1.258 - 0.006 'Isidro's new numbers
                            Return SAVI * 1.258 - 0.006 'Isidro's new numbers
                        'Return Limit(Kcint, 0.12, 0.9)
                        Case Functions.Cover.CornNLKcb
                            Dim SAVIup As Single = 0.68
                            Dim SAVIlo As Single = 0.1
                            Dim aoverb As Single = 1.11
                            Dim Kcbup As Single = 0.95
                            Return Kcbup * (1 - ((SAVIup - Math.Min(SAVIup, SAVI)) / (SAVIup - SAVIlo)) ^ aoverb)
                        Case Functions.Cover.SoybeanNLKcb
                            Dim SAVIup As Single = 0.72
                            Dim SAVIlo As Single = 0.1
                            Dim aoverb As Single = 1.21
                            Dim Kcbup As Single = 0.9
                            Return Kcbup * (1 - ((SAVIup - Math.Min(SAVIup, SAVI)) / (SAVIup - SAVIlo)) ^ aoverb)
                        Case Functions.Cover.Cotton
                            Return 0
                        Case Functions.Cover.Dryland_Cotton
                            Return 0
                        Case Functions.Cover.Sugar_Cane '@Ashish: Added Sugarcane Kcb calculation capablity
                            Return SAVI * 1.5567 - 0.0356
                        Case Else
                            Return 0
                    End Select
                Case Else 'if no reference type is chosen no kcb is returned
                    Return 0
            End Select
        ElseIf KcbVI = KcbVI.NDVI 'this is the user input Kcbrf relationship
            If InputSlope = -999 Or InputIntercept = -999 Then
                Return 0
            Else
                Return Limit(NDVI * InputSlope + InputIntercept, InputKcbMin, InputKcbMax)
            End If
        Else 'Default is SAVI 'this is the user input Kcbrf relationship
            If InputSlope = -999 Or InputIntercept = -999 Then
                Return 0
            Else
                Return Limit(SAVI * InputSlope + InputIntercept, InputKcbMin, InputKcbMax)
            End If
        End If
    End Function

    Function FindKcbVI(ByVal Cover As Cover, ByVal ETReferenceType As ETReferenceType)
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '!!!!!!  UPDATES IN THIS FUNCTION MUST BE ACCOMPANIED BY UPDATES IN FUNCTION: calcKcbReflecance() !!!!!!!
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Select Case ETReferenceType
            Case Functions.ETReferenceType.Short_Grass
                Select Case Cover
                    Case Functions.Cover.Corn, Cover.Soybean, Cover.Cotton, Cover.Dryland_Cotton
                        Return AllKcbVI.SAVI
                    Case Else
                        Return AllKcbVI.SAVI
                End Select
            Case Functions.ETReferenceType.Tall_Alfalfa 'Needs to be populated
                Select Case Cover
                    Case Functions.Cover.Corn
                        Return AllKcbVI.SAVI
                    Case Functions.Cover.Soybean
                        Return AllKcbVI.SAVI
                    Case Functions.Cover.Cotton
                        Return AllKcbVI.NDVI 'Added for Debuging
                    Case Functions.Cover.Dryland_Cotton
                        Return AllKcbVI.NDVI
                    Case Else
                        Return AllKcbVI.SAVI
                End Select
            Case Else 'if no reference type is chosen no kcb is returned
                Return 0
        End Select
    End Function


    Function calcKcbPolynomial(ByVal Cover As Cover, ByVal DOY As Integer, ByVal DOYini As Integer, ByVal DOYefc As Integer, ByVal DOYterm As Integer, ByVal ETReferenceType As ETReferenceType, ByVal KcbOffSeason As Single)
        'This Function was added by JBB
        Dim PercentIniToEFC As Single = (DOY - DOYini) / (DOYefc - DOYini) * 100
        Dim PercentEFCToTerm As Single = (DOY - DOYefc) / (DOYterm - DOYefc) * 100
        Dim KcbPoly
        Select Case ETReferenceType
            Case Functions.ETReferenceType.Short_Grass
                Select Case Cover
                    Case Functions.Cover.Corn
                        Return 0
                    Case Functions.Cover.Soybean
                        Return 0
                    Case Functions.Cover.Cotton
                        Return 0
                    Case Functions.Cover.Dryland_Cotton
                        Return 0
                    Case Else
                        Return 0
                End Select
            Case Functions.ETReferenceType.Tall_Alfalfa 'Needs to be populated
                Select Case Cover
                    Case Functions.Cover.Corn
                        If DOY >= DOYini And DOY <= DOYefc Then
                            KcbPoly = 0.1602098 - 0.001674437 * PercentIniToEFC + 0.00001334499 * PercentIniToEFC ^ 2 + 0.0000008644134 * PercentIniToEFC ^ 3 'Polynomial fit to Allen and Wright 2002 kcb by JBB
                            Return Limit(KcbPoly, 0.15, 0.96) 'Limits from Allen and Wright 2002 tabular values
                        ElseIf DOY > DOYefc And DOY <= DOYterm Then 'after EFC, Note this curve goes past harvest for Wright's original data, so Termination is not necessarily harvest date
                            KcbPoly = 1.899444 - 0.08883573 * PercentEFCToTerm + 0.002808508 * PercentEFCToTerm ^ 2 - 0.00003747604 * PercentEFCToTerm ^ 3 + 0.0000001649184 * PercentEFCToTerm ^ 4 'Polynomial fit to Allen and Wright 2002 kcb by JBB
                            Return Limit(KcbPoly, 0.15, 0.96) 'Limits from Allen and Wright 2002 tabular values
                        Else
                            Return KcbOffSeason
                        End If
                    Case Functions.Cover.Soybean
                        Return 0
                    Case Functions.Cover.Cotton
                        Return 0
                    Case Functions.Cover.Dryland_Cotton
                        Return 0
                    Case Else
                        Return 0
                End Select
            Case Else 'if no reference type is chosen no kcb is returned
                Return 0
        End Select

    End Function

    Function calcSeasonInterpolation(ByVal DoY As Integer, ByVal KcbIni As Single, ByVal KcbMid As Single, ByVal KcbEnd As Single, ByVal DateIni As Integer, ByVal DateDev As Integer, ByVal DateMid As Integer, ByVal DateLate As Integer, ByVal DateEnd As Integer)
        Select Case DoY
            Case Is > DateEnd, Is < DateIni
                Return 0
            Case Is >= DateLate
                Return KcbMid + (KcbEnd - KcbMid) * (DoY - DateLate) / (DateEnd - DateLate)
            Case Is >= DateMid
                Return KcbMid
            Case Is >= DateDev
                Return KcbIni + (KcbMid - KcbIni) * (DoY - DateDev) / (DateMid - DateDev)
            Case Is >= DateIni
                Return KcbIni
            Case Else
                Return 0
        End Select
    End Function

    Function calcKcbInterpolation(ByVal DoY As Integer, ByVal Days() As Integer, ByVal SAVIs() As Single, ByVal KcbOffSeason As Single, ByVal Cover As Cover, ByVal ETReferenceType As ETReferenceType, ByVal KcbVI As KcbVI, ByVal InputSlope As Single, ByVal InputIntercept As Single, ByVal InputKcbMax As Single, ByVal InputKcbMin As Single)
        Dim SAVIint As Single 'Interpolates SAVI and then applies Kcb, this way it isn't biased low.
        Select Case DoY
            Case Is > Days(Days.Count - 1), Is < Days(0)
                'Return 0 'commented out by JBB
                Return -999 'KcbOffSeason 'added by JBB
            Case Else
                Dim Index As Integer = -1
                If Days.Count >= 2 Then 'added by JBB
                    For I = 0 To Days.Count - 2
                        If DoY >= Days(I) And DoY < Days(I + 1) Then
                            Index = I
                            Exit For
                        End If
                    Next
                    If Index > -1 Then
                        SAVIint = SAVIs(Index) + (SAVIs(Index + 1) - SAVIs(Index)) * (DoY - Days(Index)) / (Days(Index + 1) - Days(Index))
                        Return SAVIint 'calcKcbReflectance(Cover, SAVIint, ETReferenceType, SAVIint, KcbVI, InputSlope, InputIntercept, InputKcbMax, InputKcbMin)
                    Else
                        Return SAVIs(Days.Count - 1) 'calcKcbReflectance(Cover, SAVIs(Days.Count - 1), ETReferenceType, SAVIint, KcbVI, InputSlope, InputIntercept, InputKcbMax, InputKcbMin)
                    End If
                Else ' added byJBB for single image the only Kcb is on the image date, not useful for water balance but it will provide for output of Kcb maps one at a time
                    If DoY = Days(0) Then 'added by JBB
                        'KcbOffSeason = Kcbs(0) 'added by JBB
                        Return SAVIs(0) 'calcKcbReflectance(Cover, SAVIs(0), ETReferenceType, SAVIint, KcbVI, InputSlope, InputIntercept, InputKcbMax, InputKcbMin) 'added by JBB, I believe the above line was in error, this should fix it, but HAS NOT BEEN TESTED!!!!????!!!
                    Else 'added by JBB
                        Return -999 'KcbOffSeason 'added by JBB
                    End If 'added by JBB
                End If
        End Select
    End Function

    Function calcET(ByVal G As Single, ByVal Rn As Single, ByVal LE As Single, ByVal Tair As Single, ByVal AvailableEnergy As Single, ByVal RefETinst As Single, ByVal RefET As Single, ByVal ETExtrapolation As ETExtrapolation) As Single
        'Dim Lambda As Single = (2.501 - (2.361 * 10 ^ -3) * (Tair - 273.15)) * 10 ^ 6 'See Ham 2005 Agro. Monograph Equation 5 and why is it different than LV in main.vb? JBB
        Dim Lambda As Single = 2501000 - 2361 * (Tair - 273.15) 'See Ham 2005 Agro. Monograph Equation 5 and why is it different than LV in main.vb? JBB Same as above but in form like other places in the code
        Dim ETLE As Single = 0

        Dim f = (3600 * LE / Lambda)

        Select Case ETExtrapolation
            Case ETExtrapolation.Evaporative_Fraction
                ETLE = LE / (Rn - G) * AvailableEnergy * 0.0864 / (Lambda / 1000000) 'Added by JBB, doesn't include any adjustments for advection as proposed by Gonzalez-Dugo
            Case ETExtrapolation.Reference_Evapotranspiration
                ETLE = (3600 * LE / Lambda) / RefETinst * RefET 'This is Correct JBB
        End Select

        Return Limit(ETLE, 0, 15)
    End Function

    Function ToDegrees(ByVal Value)
        Return 180 / Math.PI * Value
    End Function

    Function ToRadians(ByVal Value)
        Return Math.PI / 180 * Value
    End Function

    Function calcActualVaporPressure(ByVal SpecifiHumidity As Single, ByVal AtmosPressure As Single)
        'SpecifiHumidity and actural pressure need to be in g/kg
        ' AtmosPressure needs to be in hPa

        'convert AtmosPressure from mb to hPa (1 mb= hPa)
        AtmosPressure = AtmosPressure * 1

        'Convert SpecificHumidity from kg/kg to g/kg
        SpecifiHumidity = SpecifiHumidity * 1000

        Dim epsilon As Single = 0.622
        Dim SH = SpecifiHumidity / 1000
        Return (SH * AtmosPressure / (epsilon * (1 - SH) + SH))  'converted back to mb
    End Function

    Public Function GetNearestDateImageIndex(ByVal MultispectralImage As String, ByVal TargetImages As List(Of String)) As Integer
        Dim Index As Integer = -1

        Dim MultispectralDate As DateTime = GetDateFromPath(MultispectralImage) 'Grabs the date from the input image file name JBB
        Dim ImageDate As New List(Of DateTime) 'Creates a list of target image dates
        For I = 0 To TargetImages.Count - 1
            ImageDate.Add(GetDateFromPath(TargetImages(I))) 'Populates list of target image dates
        Next
        If ImageDate.Count = 1 Then
            Index = 0 'If only one target image, it is matched with multispectral .img. JBB
        ElseIf MultispectralDate <= ImageDate(0) Then
            Index = 0 'If multispectral .img is earlier than all target images, the earliest target is matched JBB
        ElseIf MultispectralDate >= ImageDate(TargetImages.Count - 1) Then
            Index = TargetImages.Count - 1 'If multispectral .img is later than all target images, the latest target is matched JBB
        Else
            For I = 0 To TargetImages.Count - 2 'Matches multispectral .img to target .img that is nearest in time JBB
                If MultispectralDate > ImageDate(I) And MultispectralDate <= ImageDate(I + 1) Then 'Matches multispectral .img to target .img that is nearest in time JBB
                    'If MultispectralDate - ImageDate(I) < MultispectralDate - ImageDate(0) Then 'Commented out by JBB, This logic is wierd and works differently depending on how many images are entered, how many are before the multispectral image, etc.
                    If MultispectralDate - ImageDate(I) <= ImageDate(I + 1) - MultispectralDate Then 'added by JBB, this selects the closest image, with preference for the earlier image if they are equal. 
                        Index = I 'Matches multispectral .img to target .img that is nearest in time JBB
                    Else 'Matches multispectral .img to target .img that is nearest in time JBB
                        Index = I + 1 'Matches multispectral .img to target .img that is nearest in time JBB
                    End If 'Matches multispectral .img to target .img that is nearest in time JBB
                End If
            Next
        End If

        Return Index
    End Function

    Public Function GetSameDateImageIndex(ByVal MultispectralImage As String, ByVal TargetImages As List(Of String)) As Integer
        Dim Index As Integer = -1

        Dim MultispectralDate = GetDateFromPath(MultispectralImage)
        For I = 0 To TargetImages.Count - 1
            If MultispectralDate = GetDateFromPath(TargetImages(I)) Then
                Index = I
                Exit For
            End If
        Next

        Return Index
    End Function

    Public Function GetSameDateImageIndex(ByVal RecordDate As DateTime, ByVal TargetImages As List(Of String)) As Integer
        Dim Index As Integer = -1
        '***This function is original code and will match record date as a date-time to images with date-time JBB
        For I = 0 To TargetImages.Count - 1
            If RecordDate = GetDateFromPath(TargetImages(I)) Then Index = I : Exit For
        Next

        Return Index
    End Function
    Public Function GetSameDateOnlyImageIndex(ByVal RecordDate As DateTime, ByVal TargetImages As List(Of String)) As Integer
        Dim Index As Integer = -1
        '*** This function was added by JBB because in some places RecordDate is only a date not a datetime JBB
        For I = 0 To TargetImages.Count - 1
            If RecordDate = GetDateFromPath(TargetImages(I)).Date Then Index = I : Exit For 'Adjusted code by JBB so that dates will match up JBB
        Next

        Return Index
    End Function

    Public Function GetDateFromPath(ByVal Path As String) As DateTime
        Dim DateString As String = Mid(IO.Path.GetFileNameWithoutExtension(Path), IO.Path.GetFileNameWithoutExtension(Path).Length - 15, 16)
        Return DateTime.ParseExact(DateString, "MM-dd-yyyy HH-mm", Nothing)
    End Function

    Public Function Clean(ByVal Value)
        If (Value) = Single.MinValue Then
            Dim f = 4
            f = 5

        End If
        If Value.ToString.Contains("Infinity") Then
            Return 0
        Else
            Return Value
        End If
    End Function

    Public Function CleanNull(ByVal Value)
        If IsDBNull(Value) Then
            Return 0
        Else
            Return Value
        End If
    End Function
    Public Function CleanNull_2(ByVal Value)
        Dim xx = Single.MinValue

        If (Value) < -1000 Then
            'If IsDBNull(Value) Then

            xx = Single.MinValue
            'Return 0
        Else
            Return Value
        End If
        Return xx
    End Function

    Public Function calcETAssimilation(ByRef ETcAdjusted As Single, ByRef EnergyBalanceET As Single, ByVal Method As DataAssimilation, ByVal AssimilationWeight As Single)
        Select Case Method
            Case DataAssimilation.single_Weight
                'Return ETcAdjusted + 0.78 * (EnergyBalanceET - ETcAdjusted)'commented out by JBB
                Return ETcAdjusted + AssimilationWeight * (EnergyBalanceET - ETcAdjusted) 'added by JBB
            Case Else
                Return ETcAdjusted
        End Select
    End Function


    '*********Linear Regression Functions Added by JBB ************************
    Function Rsqr(ByRef ArrX() As Single, ByRef ArrY() As Single)
        'Calculates Rsquared  following Bluman Elementary Statistics 6th ed JBB 1/5/2016

        Dim n As Integer
        Dim i As Integer
        Dim SumY As Single
        Dim SumX As Single
        Dim SumY2 As Single
        Dim SumX2 As Single
        Dim SumXY As Single
        n = ArrX.Count
        If ArrY.Count <> n Or n < 2 Then
            Rsqr = 0
            Exit Function
        End If
        SumY = 0
        SumX = 0
        SumY2 = 0
        SumX2 = 0
        For i = 0 To n - 1
            SumX = SumX + ArrX(i)
            SumY = SumY + ArrY(i)
            SumX2 = SumX2 + (ArrX(i) ^ 2)
            SumY2 = SumY2 + (ArrY(i) ^ 2)
            SumXY = SumXY + (ArrX(i) * ArrY(i))
        Next i
        Rsqr = ((n * SumXY - SumX * SumY) ^ 2) / ((n * SumX2 - SumX * SumX) * (n * SumY2 - SumY * SumY))
    End Function


    Function Slp(ByRef ArrX() As Single, ByRef ArrY() As Single)
        'Calculates Slope following Bluman Elementary Statistics 6th ed JBB 1/5/2016

        Dim n As Integer
        Dim i As Integer
        Dim SumY As Single
        Dim SumX As Single
        Dim SumX2 As Single
        Dim SumXY As Single
        n = ArrX.Count
        If ArrY.Count <> n Or n < 2 Then
            Slp = 0
            Exit Function
        End If
        SumY = 0
        SumX = 0
        SumX2 = 0
        SumXY = 0
        For i = 0 To n - 1
            SumX = SumX + ArrX(i)
            SumY = SumY + ArrY(i)
            SumX2 = SumX2 + (ArrX(i) ^ 2)
            SumXY = SumXY + (ArrX(i) * ArrY(i))
        Next i
        Slp = (n * SumXY - SumX * SumY) / (n * SumX2 - SumX * SumX)
    End Function

    Function Incpt(ByRef ArrX() As Single, ByRef ArrY() As Single)
        'Calculates Intercpet following Bluman Elementary Statistics 6th ed JBB 1/5/2016
        Dim n As Integer
        Dim i As Integer
        Dim SumY As Double
        Dim SumX As Double
        Dim SumY2 As Double
        Dim SumX2 As Double
        Dim SumXY As Double

        n = ArrX.Count
        If ArrY.Count <> n Or n < 2 Then
            Incpt = 0
            Exit Function
        End If
        SumY = 0
        SumX = 0
        SumY2 = 0
        SumXY = 0
        For i = 0 To n - 1
            SumX = SumX + ArrX(i)
            SumY = SumY + ArrY(i)
            SumX2 = SumX2 + (ArrX(i) ^ 2)
            SumXY = SumXY + (ArrX(i) * ArrY(i))
        Next i
        Incpt = (SumY * SumX2 - SumX * SumXY) / (n * SumX2 - SumX * SumX)
    End Function
    '*********End Linear Regression Functions

    '********************************************************
    'Following Schaudt and Dickinson 2000, I went to the original source, but also used PyTSEB as a resource,
    'Therefore this code may fall under the PyTSEB License see:https://github.com/hectornieto/pyTSEB
    '************************************************************
    Function SchaudtZom(ByVal Hc As Single, ByVal LAI As Single, ByVal LambdaRaupach As Single)
        Dim Zomdh As Single
        If LambdaRaupach <= 0.152 Then
            Dim a1 As Single = 5.86 'See Eq 6a Schaudt and Dickinson
            Dim b1 As Single = 10.9
            Dim c1 As Single = 1.12
            Dim d1 As Single = 1.33
            Dim Z00dh As Single = 0.00086
            Zomdh = a1 * Math.Exp(-b1 * LambdaRaupach ^ c1) * LambdaRaupach ^ d1 + Z00dh 'Eq 6a S & D 2000
        Else
            Dim a2 As Single = 0.0537 'See Eq 6a Schaudt and Dickinson
            Dim b2 As Single = 10.9
            Dim c2 As Single = 0.874
            Dim d2 As Single = 0.51
            Dim f2 As Single = 0.00368
            Zomdh = a2 * (1 - Math.Exp(-b2 * LambdaRaupach ^ c2)) / LambdaRaupach ^ d2 + f2 'Eq 6b S & D 2000
        End If
        Dim fz As Single
        If LAI < 0.8775 Then 'S&D2000 EQ 9a and Eq 9b
            fz = 0.3299 * LAI ^ 1.5 + 2.1713 'Eq 9a
        Else
            fz = 1.6771 * Math.Exp(-0.1717 * LAI) + 1 'Eq 9b
        End If
        SchaudtZom = Hc * Zomdh * fz
    End Function
    Function SchaudtD(ByVal Hc As Single, ByVal LAI As Single, ByVal LambdaRaupach As Single)
        Dim a3 As Single = 15
        Dim Ddh As Single = 1 - (1 - Math.Exp(-Math.Sqrt(a3 * LambdaRaupach))) / Math.Sqrt(a3 * LambdaRaupach) 'S&D2000 Eq 7
        Dim fd As Single = 1 - 0.3991 * Math.Exp(-0.1779 * LAI) 'Eq 10
        SchaudtD = Ddh * Hc * fd
    End Function
    '****Stability Terms add by JBB following Brutseart 2005 Hydrology and Introduction, this follows PyTSEB, therefore it may be subject
    '****to the general public license agreement of PyTSEB see:https://github.com/hectornieto/pyTSEB
    Function CalcY_MorHStable(ByVal Z As Single, ByVal d As Single, ByVal L_MO As Single) 'Caution Input Z should be Zt for Y_H calculations!!!!!!
        'Z may be Ztemp or Zwind to give y_h or y_m respectively
        Dim Zeta As Single = (Z - d) / L_MO
        Dim a1 As Single = 6.1
        Dim b1 As Single = 2.5
        'Bruseart 2005 Hydrology Eq. 2.59
        CalcY_MorHStable = -a1 * Math.Log(Zeta + (1 + Zeta ^ b1) ^ (1 / b1))
    End Function

    Function CalcY_MUnstable(ByVal Z As Single, ByVal D As Single, ByVal L_MO As Single)
        Dim Zeta As Single = (Z - D) / L_MO
        Dim y As Single = -Zeta
        Dim a1 As Single = 0.33
        Dim b1 As Single = 0.41
        Dim b3 As Single = b1 ^ (-3)
        Dim y0 As Single = -Math.Log(a1) + Math.Sqrt(3) * b1 * a1 ^ (1 / 3) * Math.PI / 6
        Dim X As Single
        'Bruseart 2005 Eq. 2.63
        If y > b3 Then y = b3
        X = (y / a1) ^ (1 / 3)
        CalcY_MUnstable = Math.Log(a1 + y) - 3 * b1 * y ^ (1 / 3) + b1 * a1 ^ (1 / 3) / 2 * Math.Log((1 + X) ^ 2 / (1 - X + X ^ 2)) + Math.Sqrt(3) * b1 * a1 ^ (1 / 3) * Math.Atan((2 * X - 1) / Math.Sqrt(3)) + y0

    End Function
    Function CalcY_HUnstable(ByVal Zt As Single, ByVal D As Single, ByVal L_MO As Single)
        Dim ZetaT As Single = (Zt - D) / L_MO
        Dim yt As Single = -ZetaT
        Dim c1 As Single = 0.33
        Dim d1 As Single = 0.057
        Dim n1 As Single = 0.78
        'Bruseart 2005 Eq. 2.64
        CalcY_HUnstable = ((1 - d1) / n1) * Math.Log((c1 + yt ^ n1) / c1)
    End Function
    '*****End added following PyTSEB*******************************************

    '************************************End Following PyTSEB*****************************************


#End Region

#Region "Classes"

    Class CoverPoint 'JBB Commented out Legacy Variables
        Public CoverName As New List(Of String)
        'Public KcbInitial As New List(Of Single)
        'Public KcbMid As New List(Of Single)
        'Public KcbEnd As New List(Of Single)
        'Public PeriodInitial As New List(Of Integer)
        'Public PeriodDevelopment As New List(Of String)
        'Public PeriodMid As New List(Of Integer)
        'Public PeriodEnd As New List(Of Integer)
        Public MaximumRootDepth As New List(Of Single)
        Public MaximumCoverHeight As New List(Of Single)
        Public MinimumCoverHeight As New List(Of Single)
        'Public DateInitial As New List(Of Integer)
        Public P As New List(Of Single)
        Public CurveNumber As New List(Of Single) 'Added by JBB
        Public PercentEffective As New List(Of Single) 'Added by JBB
        Public EvaporativeDepth As New List(Of Single) 'Added by JBB
        Public DOYStartWB As New List(Of Integer) 'Added by JBB
        Public DOYEndWB As New List(Of Integer) 'Added by JBB
        Public DOYiniMin As New List(Of Integer) 'Added by JBB
        Public DOYiniMax As New List(Of Integer) 'Added by JBB
        Public DOYefcMin As New List(Of Integer) 'Added by JBB
        Public DOYefcMax As New List(Of Integer) 'Added by JBB
        Public DOYtermMin As New List(Of Integer) 'Added by JBB
        Public DOYtermMax As New List(Of Integer) 'Added by JBB
        Public MinimumRootDepth As New List(Of Single) 'Added by JBB
        Public KcMax As New List(Of Single) 'Added by JBB
        Public KcbOffSeason As New List(Of Single) 'added by JBB
        Public YearStartWB As New List(Of Integer) 'added by JBB to allow for multiple years of running
        Public WeightForAssimilation As New List(Of Single) ' added by JBB
        Public MAD As New List(Of Single) 'added by JBB
        Public TargetDepthAboveMAD As New List(Of Single) 'added by JBB
        Public ApplicationEfficiency As New List(Of Single) 'added by JBB
        Public WeightForDepletion As New List(Of Single) ' added by JBB
        Public WeightForThetaVL As New List(Of Single) ' added by JBB
        Public WeightForEvaporatedDepth As New List(Of Single) ' added by JBB
        Public FalseEndSAVI As New List(Of Single) ' added by JBB
        Public FalseEndSAVIDOY As New List(Of Integer) ' added by JBB
        Public FalsePeakSAVI As New List(Of Single) ' added by JBB
        Public FalsePeakSAVIDOY As New List(Of Integer) ' added by JBB
        Public FalseEndSAVIGDD As New List(Of Single) 'added by JBB 2/1/2018
        Public FalsePeakSAVIGDD As New List(Of Single) 'added by JBB 2/1/2018
        Public GDDBase As New List(Of Single) 'added by JBB 2/1/2018
        Public GDDMaxTemp As New List(Of Single) 'added by JBB 2/1/2018
        Public WeightForSkinEvap As New List(Of Single) 'added by JBB 2/9/2018 copying others.
        Public RatioKcbLate As New List(Of Single) 'added by JBB 2/16/2018 copying others
        Public StressBaseTemp As New List(Of Single) 'added by JBB 2/16/2018 copying others
        Public StressMaxTemp As New List(Of Single) 'added by JBB 2/16/2018 copying others
        Public xBiomassMaxSAVI As New List(Of Single) 'added by JBB 8/17/2018 copying others

        Sub Populate(ByVal CoverTable As ESRI.ArcGIS.Geodatabase.ITable, ByVal CoverPointIndex As CoverPointIndex)
            CoverName.Clear()
            'KcbInitial.Clear()
            'KcbMid.Clear()
            'KcbEnd.Clear()
            'PeriodInitial.Clear()
            'PeriodDevelopment.Clear()
            'PeriodMid.Clear()
            'PeriodEnd.Clear()
            MaximumRootDepth.Clear()
            MaximumCoverHeight.Clear()
            MinimumCoverHeight.Clear()
            'DateInitial.Clear()
            P.Clear()
            CurveNumber.Clear() 'Added by JBB
            PercentEffective.Clear() 'Added by JBB
            EvaporativeDepth.Clear() 'Added by JBB
            DOYStartWB.Clear() 'Added by JBB
            DOYEndWB.Clear() 'Added by JBB
            DOYiniMin.Clear() 'Added by JBB
            DOYiniMax.Clear() 'Added by JBB
            DOYefcMin.Clear() 'Added by JBB
            DOYefcMax.Clear() 'Added by JBB
            DOYtermMin.Clear() 'Added by JBB
            DOYtermMax.Clear() 'Added by JBB
            MinimumRootDepth.Clear() 'Added by JBB
            KcMax.Clear() 'Added by JBB
            KcbOffSeason.Clear() 'added by JBB
            YearStartWB.Clear() 'added byJBB
            WeightForAssimilation.Clear() 'added by JBB
            MAD.Clear() 'added by JBB
            TargetDepthAboveMAD.Clear() 'added by JBB
            ApplicationEfficiency.Clear() 'added by JBB
            WeightForDepletion.Clear() 'added by JBB
            WeightForThetaVL.Clear() 'added by JBB
            WeightForEvaporatedDepth.Clear() 'added by JBB
            FalseEndSAVI.Clear() 'added by JBB
            FalseEndSAVIDOY.Clear() 'added by JBB
            FalsePeakSAVI.Clear() 'added by JBB
            FalsePeakSAVIDOY.Clear() 'added by JBB
            FalseEndSAVIGDD.Clear() 'added by JBB 2/1/2018
            FalsePeakSAVIGDD.Clear() 'added by JBB 2/1/2018
            GDDBase.Clear() 'added by JBB 2/1/2018
            GDDMaxTemp.Clear() 'added by JBB 2/1/2018
            WeightForSkinEvap.Clear() 'added by JBB 2/9/2018 copying others.
            RatioKcbLate.Clear() 'added by JBB 2/16/2018 copying others.
            StressBaseTemp.Clear() 'added by JBB 2/16/2018 copying others.
            StressMaxTemp.Clear() 'added by JBB 2/16/2018 copying others.
            xBiomassMaxSAVI.Clear() 'added by JBB 8/17/2018 copying others.
            '**************ADDED BY JBB AS A Reminder
            'MsgBox("The Function CoverPoint.Populate is hard coded for two covers, change as necessary JBB") 'See the If Count = 2 Then Exit Do line below JBB This problem was fixed by JBB

            '******************End ADDED by JBB
            Dim CoverCursor As ESRI.ArcGIS.Geodatabase.ICursor = CoverTable.Search(Nothing, True)
            Dim CoverRow As ESRI.ArcGIS.Geodatabase.IRow = CoverCursor.NextRow()
            Dim Count As Integer = 0
            Do While Not CoverRow Is Nothing
                'If Count = 2 Then Exit Do 'added by JBB, must be incremented to total number of rows in cover table
                For MM = 1 To CoverRow.Fields.FieldCount - 1 'added by JBB
                    If CoverRow.Value(MM).ToString = "" Then Exit Do 'added by JBB to exit the function if a row of blank values is next in the table
                Next MM
                CoverName.Add(0) : If CoverPointIndex.CoverName > -1 Then CoverName(Count) = CoverRow.Value(CoverPointIndex.CoverName)
                'KcbInitial.Add(0) : If CoverPointIndex.KcbInitial > -1 Then KcbInitial(Count) = CoverRow.Value(CoverPointIndex.KcbInitial)
                'KcbMid.Add(0) : If CoverPointIndex.KcbMid > -1 Then KcbMid(Count) = CoverRow.Value(CoverPointIndex.KcbMid)
                'KcbEnd.Add(0) : If CoverPointIndex.KcbEnd > -1 Then KcbEnd(Count) = CoverRow.Value(CoverPointIndex.KcbEnd)
                'PeriodInitial.Add(0) : If CoverPointIndex.PeriodInitial > -1 Then PeriodInitial(Count) = CoverRow.Value(CoverPointIndex.PeriodInitial)
                'PeriodDevelopment.Add(0) : If CoverPointIndex.PeriodDevelopment > -1 Then PeriodDevelopment(Count) = CoverRow.Value(CoverPointIndex.PeriodDevelopment)
                'PeriodMid.Add(0) : If CoverPointIndex.PeriodMid > -1 Then PeriodMid(Count) = CoverRow.Value(CoverPointIndex.PeriodMid)
                'PeriodEnd.Add(0) : If CoverPointIndex.PeriodEnd > -1 Then PeriodEnd(Count) = CoverRow.Value(CoverPointIndex.PeriodEnd)
                MaximumRootDepth.Add(0) : If CoverPointIndex.MaximumRootDepth > -1 Then MaximumRootDepth(Count) = CoverRow.Value(CoverPointIndex.MaximumRootDepth)
                MaximumCoverHeight.Add(0) : If CoverPointIndex.MaximumCoverHeight > -1 Then MaximumCoverHeight(Count) = CoverRow.Value(CoverPointIndex.MaximumCoverHeight)
                MinimumCoverHeight.Add(0) : If CoverPointIndex.MinimumCoverHeight > -1 Then MinimumCoverHeight(Count) = CoverRow.Value(CoverPointIndex.MinimumCoverHeight)
                'DateInitial.Add(0) : If CoverPointIndex.DateInitial > -1 Then DateInitial(Count) = CoverRow.Value(CoverPointIndex.DateInitial)
                P.Add(0) : If CoverPointIndex.P > -1 Then P(Count) = CoverRow.Value(CoverPointIndex.P)
                CurveNumber.Add(0) : If CoverPointIndex.CurveNumber > -1 Then CurveNumber(Count) = CoverRow.Value(CoverPointIndex.CurveNumber) 'Added by JBB
                PercentEffective.Add(0) : If CoverPointIndex.PercentEffective > -1 Then PercentEffective(Count) = CoverRow.Value(CoverPointIndex.PercentEffective) 'Added by JBB
                EvaporativeDepth.Add(0) : If CoverPointIndex.EvaporativeDepth > -1 Then EvaporativeDepth(Count) = CoverRow.Value(CoverPointIndex.EvaporativeDepth) 'Added by JBB
                DOYStartWB.Add(0) : If CoverPointIndex.DOYStartWB > -1 Then DOYStartWB(Count) = CoverRow.Value(CoverPointIndex.DOYStartWB) 'Added by JBB
                DOYEndWB.Add(0) : If CoverPointIndex.DOYEndWB > -1 Then DOYEndWB(Count) = CoverRow.Value(CoverPointIndex.DOYEndWB) 'Added by JBB
                DOYiniMin.Add(0) : If CoverPointIndex.DOYiniMin > -1 Then DOYiniMin(Count) = CoverRow.Value(CoverPointIndex.DOYiniMin) 'Added by JBB
                DOYiniMax.Add(0) : If CoverPointIndex.DOYiniMax > -1 Then DOYiniMax(Count) = CoverRow.Value(CoverPointIndex.DOYiniMax) 'Added by JBB
                DOYefcMax.Add(0) : If CoverPointIndex.DOYefcMax > -1 Then DOYefcMax(Count) = CoverRow.Value(CoverPointIndex.DOYefcMax) 'Added by JBB
                DOYefcMin.Add(0) : If CoverPointIndex.DOYefcMin > -1 Then DOYefcMin(Count) = CoverRow.Value(CoverPointIndex.DOYefcMin) 'Added by JBB
                DOYtermMax.Add(0) : If CoverPointIndex.DOYtermMax > -1 Then DOYtermMax(Count) = CoverRow.Value(CoverPointIndex.DOYtermMax) 'Added by JBB
                DOYtermMin.Add(0) : If CoverPointIndex.DOYtermMin > -1 Then DOYtermMin(Count) = CoverRow.Value(CoverPointIndex.DOYtermMin) 'Added by JBB
                MinimumRootDepth.Add(0) : If CoverPointIndex.MinimumRootDepth > -1 Then MinimumRootDepth(Count) = CoverRow.Value(CoverPointIndex.MinimumRootDepth) 'Added by JBB
                KcMax.Add(0) : If CoverPointIndex.KcMax > -1 Then KcMax(Count) = CoverRow.Value(CoverPointIndex.KcMax) 'Added by JBB
                KcbOffSeason.Add(0) : If CoverPointIndex.KcbOffSeason > -1 Then KcbOffSeason(Count) = CoverRow.Value(CoverPointIndex.KcbOffSeason) 'Added by JBB
                YearStartWB.Add(0) : If CoverPointIndex.YearStartWB > -1 Then YearStartWB(Count) = CoverRow.Value(CoverPointIndex.YearStartWB) 'added by JBB
                WeightForAssimilation.Add(0) : If CoverPointIndex.WeightForAssimilation > -1 Then WeightForAssimilation(Count) = CoverRow.Value(CoverPointIndex.WeightForAssimilation) 'Added by JBB
                MAD.Add(0) : If CoverPointIndex.MAD > -1 Then MAD(Count) = CoverRow.Value(CoverPointIndex.MAD) 'Added by JBB
                TargetDepthAboveMAD.Add(0) : If CoverPointIndex.TargetDepthAboveMAD > -1 Then TargetDepthAboveMAD(Count) = CoverRow.Value(CoverPointIndex.TargetDepthAboveMAD) 'Added by JBB
                ApplicationEfficiency.Add(0) : If CoverPointIndex.ApplicationEfficiency > -1 Then ApplicationEfficiency(Count) = CoverRow.Value(CoverPointIndex.ApplicationEfficiency) 'Added by JBB
                WeightForDepletion.Add(0) : If CoverPointIndex.WeightForDepletion > -1 Then WeightForDepletion(Count) = CoverRow.Value(CoverPointIndex.WeightForDepletion) 'Added by JBB
                WeightForThetaVL.Add(0) : If CoverPointIndex.WeightForThetaVL > -1 Then WeightForThetaVL(Count) = CoverRow.Value(CoverPointIndex.WeightForThetaVL) 'Added by JBB
                WeightForEvaporatedDepth.Add(0) : If CoverPointIndex.WeightForEvaporatedDepth > -1 Then WeightForEvaporatedDepth(Count) = CoverRow.Value(CoverPointIndex.WeightForEvaporatedDepth) 'Added by JBB
                FalseEndSAVI.Add(0) : If CoverPointIndex.FalseEndSAVI > -1 Then FalseEndSAVI(Count) = CoverRow.Value(CoverPointIndex.FalseEndSAVI) 'Added by JBB
                FalseEndSAVIDOY.Add(0) : If CoverPointIndex.FalseEndSAVIDOY > -1 Then FalseEndSAVIDOY(Count) = CoverRow.Value(CoverPointIndex.FalseEndSAVIDOY) 'Added by JBB
                FalsePeakSAVI.Add(0) : If CoverPointIndex.FalsePeakSAVI > -1 Then FalsePeakSAVI(Count) = CoverRow.Value(CoverPointIndex.FalsePeakSAVI) 'Added by JBB
                FalsePeakSAVIDOY.Add(0) : If CoverPointIndex.FalsePeakSAVIDOY > -1 Then FalsePeakSAVIDOY(Count) = CoverRow.Value(CoverPointIndex.FalsePeakSAVIDOY) 'Added by JBB
                FalseEndSAVIGDD.Add(0) : If CoverPointIndex.FalseEndSAVIGDD > -1 Then FalseEndSAVIGDD(Count) = CoverRow.Value(CoverPointIndex.FalseEndSAVIGDD) 'Added by JBB
                FalsePeakSAVIGDD.Add(0) : If CoverPointIndex.FalsePeakSAVIGDD > -1 Then FalsePeakSAVIGDD(Count) = CoverRow.Value(CoverPointIndex.FalsePeakSAVIGDD) 'Added by JBB
                GDDBase.Add(0) : If CoverPointIndex.GDDBase > -1 Then GDDBase(Count) = CoverRow.Value(CoverPointIndex.GDDBase) 'Added by JBB
                GDDMaxTemp.Add(0) : If CoverPointIndex.GDDMaxTemp > -1 Then GDDMaxTemp(Count) = CoverRow.Value(CoverPointIndex.GDDMaxTemp) 'Added by JBB
                WeightForSkinEvap.Add(0) : If CoverPointIndex.WeightForSkinEvap > -1 Then WeightForSkinEvap(Count) = CoverRow.Value(CoverPointIndex.WeightForSkinEvap) 'Added by JBB
                RatioKcbLate.Add(0) : If CoverPointIndex.RatioKcbLate > -1 Then RatioKcbLate(Count) = CoverRow.Value(CoverPointIndex.RatioKcbLate) ''added by JBB 2/16/2018 copying others.
                StressBaseTemp.Add(0) : If CoverPointIndex.StressBaseTemp > -1 Then StressBaseTemp(Count) = CoverRow.Value(CoverPointIndex.StressBaseTemp) 'added by JBB 2/16/2018 copying others.
                StressMaxTemp.Add(0) : If CoverPointIndex.StressMaxTemp > -1 Then StressMaxTemp(Count) = CoverRow.Value(CoverPointIndex.StressMaxTemp) 'added by JBB 2/16/2018 copying others.
                xBiomassMaxSAVI.Add(0) : If CoverPointIndex.xBiomassMaxSavi > -1 Then xBiomassMaxSAVI(Count) = CoverRow.Value(CoverPointIndex.xBiomassMaxSavi) 'added by JBB 8/17/2018 copying others.
                Count += 1
                CoverRow = CoverCursor.NextRow()
            Loop
            System.Runtime.InteropServices.Marshal.ReleaseComObject(CoverCursor)
        End Sub

    End Class

    Class CoverPointIndex
        Public CoverName As Integer
        'Public KcbInitial As Integer
        'Public KcbMid As Integer
        'Public KcbEnd As Integer
        'Public PeriodInitial As Integer
        'Public PeriodDevelopment As Integer
        'Public PeriodMid As Integer
        'Public PeriodEnd As Integer
        Public MaximumRootDepth As Integer
        Public MaximumCoverHeight As Integer
        Public MinimumCoverHeight As Integer
        'Public DateInitial As Integer
        Public P As Integer
        Public CurveNumber As Integer 'Added by JBB
        Public PercentEffective As Integer 'Added by JBB
        Public EvaporativeDepth As Integer 'Added by JBB
        Public DOYStartWB As Integer 'Added by JBB
        Public DOYEndWB As Integer 'Added by JBB
        Public DOYiniMin As Integer 'Added by JBB
        Public DOYiniMax As Integer 'Added by JBB
        Public DOYefcMin As Integer 'Added by JBB
        Public DOYefcMax As Integer 'Added by JBB
        Public DOYtermMin As Integer 'Added by JBB
        Public DOYtermMax As Integer 'Added by JBB
        Public MinimumRootDepth As Integer 'Added by JBB
        Public KcMax As Integer 'Added by JBB
        Public KcbOffSeason As Integer 'added by JBB
        Public YearStartWB As Integer 'added by JBB
        Public WeightForAssimilation As Integer 'added by JBB
        Public MAD As Integer 'added by JBB
        Public TargetDepthAboveMAD As Integer 'added by JBB
        Public ApplicationEfficiency As Integer 'added by JBB
        Public WeightForDepletion As Integer 'added by JBB
        Public WeightForThetaVL As Integer 'added by JBB
        Public WeightForEvaporatedDepth As Integer 'added by JBB
        Public FalseEndSAVI As Integer 'added by JBB
        Public FalseEndSAVIDOY As Integer 'added by JBB
        Public FalsePeakSAVI As Integer 'added by JBB
        Public FalsePeakSAVIDOY As Integer 'added by JBB
        Public FalseEndSAVIGDD As Integer 'added by JBB 2/1/2018
        Public FalsePeakSAVIGDD As Integer 'added by JBB 2/1/2018
        Public GDDBase As Integer ' added by JBB 2/1/2018
        Public GDDMaxTemp As Integer 'added by JBB 2/1/2018
        Public WeightForSkinEvap As Integer 'added by JBB 2/9/2018 following others
        Public RatioKcbLate As Integer 'added by JBB 2/16/2018 copying others.
        Public StressBaseTemp As Integer 'added by JBB 2/16/2018 copying others.
        Public StressMaxTemp As Integer 'added by JBB 2/16/2018 copying others.
        Public xBiomassMaxSavi As Integer 'added by JBB 8/17/2018 copying others.


        Public Sub Initialize()
            CoverName = -1
            'KcbInitial = -1
            'KcbMid = -1
            'KcbEnd = -1
            'PeriodInitial = -1
            'PeriodDevelopment = -1
            'PeriodMid = -1
            'PeriodEnd = -1
            MaximumRootDepth = -1
            MaximumCoverHeight = -1
            MinimumCoverHeight = -1
            'DateInitial = -1
            P = -1
            CurveNumber = -1 'Added by JBB
            PercentEffective = -1 'Added by JBB
            EvaporativeDepth = -1 'Added by JBB
            DOYStartWB = -1 'Added by JBB
            DOYEndWB = -1 'Added by JBB
            DOYiniMin = -1 'Added by JBB
            DOYiniMax = -1 'Added by JBB
            DOYefcMin = -1 'Added by JBB
            DOYefcMax = -1 'Added by JBB
            DOYtermMin = -1 'Added by JBB
            DOYtermMax = -1 'Added by JBB
            MinimumRootDepth = -1 'Added by JBB
            KcMax = -1 'Added by JBB
            KcbOffSeason = -1 'Added by JBB
            YearStartWB = -1 'added by JBB
            WeightForAssimilation = -1 ' added by JBB
            MAD = -1 ' added by JBB
            TargetDepthAboveMAD = -1 ' added by JBB
            ApplicationEfficiency = -1 ' added by JBB
            WeightForDepletion = -1 'added by JBB
            WeightForThetaVL = -1 'added by JBB
            WeightForEvaporatedDepth = -1 'added by JBB
            FalseEndSAVI = -1 'added by JBB
            FalseEndSAVIDOY = -1 'added by JBB
            FalsePeakSAVI = -1 'added by JBB
            FalsePeakSAVIDOY = -1 'added by JBB
            FalseEndSAVIGDD = -1 'added by JBB
            FalsePeakSAVIGDD = -1 'added by JBB
            GDDBase = -1 'added by JBB
            GDDMaxTemp = -1 'added by JBB
            WeightForSkinEvap = -1 'added by JBB copying others 2/9/2018
            RatioKcbLate = -1 'added by JBB 2/16/2018 copying others.
            StressBaseTemp = -1 'added by JBB 2/16/2018 copying others.
            StressMaxTemp = -1 'added by JBB 2/16/2018 copying others.
            xBiomassMaxSavi = -1 'added by JBB 8/17/2018 copying others.
        End Sub

    End Class

    Class WBOutputDateIndex 'This class was added by JBB to handle the new water balance output date functionality JBB
        Public WBOutDate As Integer

        Public Sub Initialize()
            WBOutDate = -1
        End Sub

    End Class
    Class WBOutputDates 'Added by JBB to allow for input of a list of output dates for wb images JBB
        Public Shared WtrBalOutputDates As New List(Of Date)
        Function CountRows(ByVal WBOutputDateTbl As ESRI.ArcGIS.Geodatabase.ITable, ByVal WBOutputDateIndex As WBOutputDateIndex)
            Dim WBOutCursor As ESRI.ArcGIS.Geodatabase.ICursor = WBOutputDateTbl.Search(Nothing, True)
            Dim WBOutRow As ESRI.ArcGIS.Geodatabase.IRow = WBOutCursor.NextRow()
            Dim Count As Integer = 0
            Do While Not WBOutRow Is Nothing
                If Not IsDBNull(WBOutRow.Value(WBOutputDateIndex.WBOutDate)) Then
                    Count += 1
                End If
                WBOutRow = WBOutCursor.NextRow()
            Loop
            System.Runtime.InteropServices.Marshal.ReleaseComObject(WBOutCursor)
            CountRows = Count
        End Function

    End Class


    Class WBPointOutputIndex 'This class was added by JBB to handle the new water balance output date functionality Copied from similar code for cover output date grid which is probably copied from the cover properties grid and  JBB
        Public Xcoord As Integer
        Public Ycoord As Integer

        Public Sub Initialize()
            Xcoord = -1
            Ycoord = -1
        End Sub

    End Class
    Class WBPointOutputs 'Added by JBB to allow for input of a list of output dates for wb images JBB
        Public Shared WBPointOutputs As New List(Of Double) 'Declared as double as is the manually selected points list in Main JBB
        Function CountRows(ByVal WBPntOutTbl As ESRI.ArcGIS.Geodatabase.ITable, ByVal WBPointOutputIndex As WBPointOutputIndex)
            Dim WBPntOutCursor As ESRI.ArcGIS.Geodatabase.ICursor = WBPntOutTbl.Search(Nothing, True)
            Dim WBPntOutRow As ESRI.ArcGIS.Geodatabase.IRow = WBPntOutCursor.NextRow()
            Dim Count As Integer = 0
            Do While Not WBPntOutRow Is Nothing
                If Not IsDBNull(WBPntOutRow.Value(WBPointOutputIndex.Xcoord)) And Not IsDBNull(WBPntOutRow.Value(WBPointOutputIndex.Ycoord)) Then
                    Count += 1
                End If
                WBPntOutRow = WBPntOutCursor.NextRow()
            Loop
            System.Runtime.InteropServices.Marshal.ReleaseComObject(WBPntOutCursor)
            CountRows = Count
        End Function

    End Class



    Class WeatherPoint
        Public Shared ActualVaporPressure As New List(Of Single)
        Public Shared AirTemperature As New List(Of Single)
        Public Shared AnemometerReferenceHeight As New List(Of Single)
        Public Shared AirTemperatureReferenceHeight As New List(Of Single)
        Public Shared AtmosphericPressure As New List(Of Single)
        Public Shared CoverName As New List(Of String)
        Public Shared CoverHeight As New List(Of String) 'Added by JBB
        Public Shared Irrigation As New List(Of Single)
        Public Shared RelativeHumidity As New List(Of Single)
        Public Shared Precipitation As New List(Of Single)
        Public Shared ETDailyReference As New List(Of Single)
        Public Shared ETInstantaneous As New List(Of Single)
        Public Shared SolarRadiation As New List(Of Single)
        Public Shared WindSpeed As New List(Of Single)
        Public Shared RecordDate As New List(Of DateTime)
        Public Shared WindFetchHeight As New List(Of Single) 'Added by JBB
        Public Shared xEvaporationScalingFactor As New List(Of Single) 'Added by JBB
        Public Shared DailyAvailableEnergy As New List(Of Single) 'Added by JBB
        Public Shared DailyMaximumTemperature As New List(Of Single) 'Added by JBB
        Public Shared DailyMinimumTemperature As New List(Of Single) 'Added by JBB
        Public Shared xFetchDistanceAnemometer As New List(Of Single) 'added by JBB for wind adjust
        Public Shared xFetchDistanceCover As New List(Of Single) 'added by JBB for wind adjust
        Public Shared xIBLHeightOptional As New List(Of Single) 'added by JBB for wind adjust
        Public Shared xRegionalCoverHeight As New List(Of Single) 'added by JBB for wind adjust
        Public Shared xTypWindSpeed2m As New List(Of Single) ' added by JBB
        Public Shared xMonthlyETo As New List(Of Single) 'added by JBB copying others
        Public Shared xFractionWetToday As New List(Of Single) 'added by JBB copying others

        Sub Populate(ByVal WeatherTable As ESRI.ArcGIS.Geodatabase.ITable, ByVal WeatherPointIndex As WeatherPointIndex, ByVal RecordDate As DateTime)
            ActualVaporPressure.Clear()
            AirTemperature.Clear()
            AnemometerReferenceHeight.Clear()
            AirTemperatureReferenceHeight.Clear()
            AtmosphericPressure.Clear()
            CoverName.Clear()
            CoverHeight.Clear()
            Irrigation.Clear()
            RelativeHumidity.Clear()
            Precipitation.Clear()
            ETDailyReference.Clear()
            ETInstantaneous.Clear()
            SolarRadiation.Clear()
            WindSpeed.Clear()
            WeatherPoint.RecordDate.Clear()
            WindFetchHeight.Clear() 'added by JBB
            xEvaporationScalingFactor.Clear() 'Added by JBB
            DailyAvailableEnergy.Clear() ' added by JBB
            DailyMaximumTemperature.Clear() 'added by JBB
            DailyMinimumTemperature.Clear() 'added by JBB
            xFetchDistanceAnemometer.Clear() 'added by JBB for wind adjust
            xFetchDistanceCover.Clear() 'added by JBB for wind adjust
            xIBLHeightOptional.Clear() 'added by JBB for wind adjust
            xRegionalCoverHeight.Clear() 'added by JBB for wind adjust
            xTypWindSpeed2m.Clear() 'added by JBB for ETo Kc adjustments
            xMonthlyETo.Clear() 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
            xFractionWetToday.Clear() 'added by JBB for  ASCE MOP 70 ed2,2016 eq. 9.28 copying others

            Dim WeatherCursor As ESRI.ArcGIS.Geodatabase.ICursor = WeatherTable.Search(Nothing, True)
            Dim WeatherRow As ESRI.ArcGIS.Geodatabase.IRow = WeatherCursor.NextRow()
            Dim Count As Integer = 0
            Do While Not WeatherRow Is Nothing
                If Not IsDBNull(WeatherRow.Value(WeatherPointIndex.RecordDate)) Then
                    If CDate(WeatherRow.Value(WeatherPointIndex.RecordDate)) = RecordDate.Date Then
                        ActualVaporPressure.Add(0) : If WeatherPointIndex.ActualVaporPressure > -1 Then ActualVaporPressure(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.ActualVaporPressure))
                        AirTemperature.Add(0) : If WeatherPointIndex.AirTemperature > -1 Then AirTemperature(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.AirTemperature))
                        AnemometerReferenceHeight.Add(0)
                        If WeatherPointIndex.AnemometerReferenceHeight > -1 Then
                            AnemometerReferenceHeight(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.AnemometerReferenceHeight))
                        End If
                        AirTemperatureReferenceHeight.Add(0) : If WeatherPointIndex.AirTemperatureReferenceHeight > -1 Then AirTemperatureReferenceHeight(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.AirTemperatureReferenceHeight))
                        AtmosphericPressure.Add(0) : If WeatherPointIndex.AtmosphericPressure > -1 Then AtmosphericPressure(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.AtmosphericPressure))
                        CoverName.Add(0) : If WeatherPointIndex.CoverName > -1 Then CoverName(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.CoverName))
                        CoverHeight.Add(0) : If WeatherPointIndex.CoverHeight > -1 Then CoverHeight(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.CoverHeight))
                        ETInstantaneous.Add(0) : If WeatherPointIndex.InstantaneousShortReferenceEvapotranspiration > -1 Then ETInstantaneous(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.InstantaneousShortReferenceEvapotranspiration))
                        Irrigation.Add(0) : If WeatherPointIndex.Irrigation > -1 Then Irrigation(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.Irrigation))
                        RelativeHumidity.Add(0) : If WeatherPointIndex.RelativeHumidity > -1 Then CleanNull(RelativeHumidity(Count) = WeatherRow.Value(WeatherPointIndex.RelativeHumidity))
                        Precipitation.Add(0) : If WeatherPointIndex.Precipitation > -1 Then Precipitation(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.Precipitation))
                        ETDailyReference.Add(0) : If WeatherPointIndex.ShortReferenceEvapotranspiration > -1 Then ETDailyReference(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.ShortReferenceEvapotranspiration))
                        SolarRadiation.Add(0) : If WeatherPointIndex.SolarRadiation > -1 Then SolarRadiation(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.SolarRadiation))
                        WindSpeed.Add(0) : If WeatherPointIndex.WindSpeed > -1 Then WindSpeed(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.WindSpeed))
                        WindFetchHeight.Add(0) : If WeatherPointIndex.WindFetchHeight > -1 Then WindFetchHeight(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.WindFetchHeight)) 'Added by JBB
                        xEvaporationScalingFactor.Add(0) : If WeatherPointIndex.xEvaporationScalingFactor > -1 Then xEvaporationScalingFactor(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xEvaporationScalingFactor)) 'Added by JBB
                        DailyAvailableEnergy.Add(0) : If WeatherPointIndex.DailyAvailableEnergy > -1 Then DailyAvailableEnergy(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.DailyAvailableEnergy)) 'Added by JBB
                        DailyMaximumTemperature.Add(0) : If WeatherPointIndex.DailyMaximumTemperature > -1 Then DailyMaximumTemperature(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.DailyMaximumTemperature)) ' Added by JBB
                        DailyMinimumTemperature.Add(0) : If WeatherPointIndex.DailyMinimumTemperature > -1 Then DailyMinimumTemperature(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.DailyMinimumTemperature)) ' Added by JBB
                        xFetchDistanceAnemometer.Add(0) : If WeatherPointIndex.xFetchDistanceAnemometer > -1 Then xFetchDistanceAnemometer(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xFetchDistanceAnemometer)) 'added by JBB for wind adjust
                        xFetchDistanceCover.Add(0) : If WeatherPointIndex.xFetchDistanceCover > -1 Then xFetchDistanceCover(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xFetchDistanceCover)) 'added by JBB for wind adjust
                        xIBLHeightOptional.Add(0) : If WeatherPointIndex.xIBLHeightOptional > -1 Then xIBLHeightOptional(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xIBLHeightOptional)) 'added by JBB for wind adjust
                        xRegionalCoverHeight.Add(0) : If WeatherPointIndex.xRegionalCoverHeight > -1 Then xRegionalCoverHeight(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xRegionalCoverHeight)) 'added by JBB for wind adjust
                        xTypWindSpeed2m.Add(0) : If WeatherPointIndex.xTypWindSpeed2m > -1 Then xTypWindSpeed2m(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xTypWindSpeed2m)) 'Added by JBB
                        xMonthlyETo.Add(0) : If WeatherPointIndex.xMonthlyETo > -1 Then xMonthlyETo(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xMonthlyETo)) 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
                        xFractionWetToday.Add(0) : If WeatherPointIndex.xFractionWetToday > -1 Then xFractionWetToday(Count) = CleanNull(WeatherRow.Value(WeatherPointIndex.xFractionWetToday)) 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
                        Count += 1
                    End If
                    End If
                WeatherRow = WeatherCursor.NextRow()
            Loop
            System.Runtime.InteropServices.Marshal.ReleaseComObject(WeatherCursor)
        End Sub

        Sub Populate(ByVal WeatherTable As ESRI.ArcGIS.Geodatabase.ITable, ByVal WeatherPointIndex As WeatherPointIndex, ByVal Year As Integer)
            ActualVaporPressure.Clear()
            AirTemperature.Clear()
            AnemometerReferenceHeight.Clear()
            AirTemperatureReferenceHeight.Clear()
            AtmosphericPressure.Clear()
            WeatherPoint.CoverName.Clear()
            Irrigation.Clear()
            RelativeHumidity.Clear()
            Precipitation.Clear()
            ETDailyReference.Clear()
            ETInstantaneous.Clear()
            SolarRadiation.Clear()
            WindSpeed.Clear()
            RecordDate.Clear()
            WindFetchHeight.Clear() 'Added by JBB
            xEvaporationScalingFactor.Clear() 'Added by JBB
            DailyAvailableEnergy.Clear() 'added by JBB
            CoverHeight.Clear() 'added by JBB
            DailyMaximumTemperature.Clear() 'added by JBB
            DailyMinimumTemperature.Clear() 'added by JBB
            xFetchDistanceAnemometer.Clear() 'added by JBB for wind adjust
            xFetchDistanceCover.Clear() 'added by JBB for wind adjust
            xIBLHeightOptional.Clear() 'added by JBB for wind adjust
            xRegionalCoverHeight.Clear() 'added by JBB for wind adjust
            xTypWindSpeed2m.Clear() ' added by JBB
            xMonthlyETo.Clear() 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
            xFractionWetToday.Clear() 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others

            Dim WeatherCursor As ESRI.ArcGIS.Geodatabase.ICursor = WeatherTable.Search(Nothing, True)
            Dim WeatherRow As ESRI.ArcGIS.Geodatabase.IRow = WeatherCursor.NextRow()
            Dim Count As Integer = 0
            Do While Not WeatherRow Is Nothing
                Try
                    If Not IsDBNull(WeatherRow.Value(WeatherPointIndex.RecordDate)) Then
                        Dim DateStamp As DateTime = CDate(WeatherRow.Value(WeatherPointIndex.RecordDate))
                        'If DateStamp.Year = Year And Not RecordDate.Contains(DateStamp) Then
                        If DateStamp.Year = Year Or DateStamp.Year = Year + 1 Then 'Modified by JBB to handle two-year runs
                            If Not RecordDate.Contains(DateStamp) Then 'added by JBB
                                AnemometerReferenceHeight.Add(0) : If WeatherPointIndex.AnemometerReferenceHeight > -1 Then AnemometerReferenceHeight(Count) = WeatherRow.Value(WeatherPointIndex.AnemometerReferenceHeight)
                                Irrigation.Add(0) : If WeatherPointIndex.Irrigation > -1 Then Irrigation(Count) = WeatherRow.Value(WeatherPointIndex.Irrigation)
                                'RelativeHumidity.Add(0) : If WeatherPointIndex.Irrigation > -1 Then RelativeHumidity(Count) = WeatherRow.Value(WeatherPointIndex.RelativeHumidity)'Original Code commented out by JBB had .irrigation not .RelativeHumidity JBB
                                RelativeHumidity.Add(0) : If WeatherPointIndex.RelativeHumidity > -1 Then RelativeHumidity(Count) = WeatherRow.Value(WeatherPointIndex.RelativeHumidity) 'Corrected code by JBB
                                Precipitation.Add(0) : If WeatherPointIndex.Precipitation > -1 Then Precipitation(Count) = WeatherRow.Value(WeatherPointIndex.Precipitation)
                                ETDailyReference.Add(0) : If WeatherPointIndex.ShortReferenceEvapotranspiration > -1 Then ETDailyReference(Count) = WeatherRow.Value(WeatherPointIndex.ShortReferenceEvapotranspiration)
                                'WindSpeed.Add(0) : If WeatherPointIndex.WindSpeed > -1 Then WindSpeed(Count) = WeatherRow.Value(WeatherPointIndex.WindSpeed)
                                RecordDate.Add(DateStamp)
                                'WindFetchHeight.Add(0) : If WeatherPointIndex.WindFetchHeight > -1 Then WindFetchHeight(Count) = WeatherRow.Value(WeatherPointIndex.WindFetchHeight) 'Added by JBB
                                ' DailyAvailableEnergy.Add(0) : If WeatherPointIndex.DailyAvailableEnergy > -1 Then DailyAvailableEnergy(Count) = WeatherRow.Value(WeatherPointIndex.DailyAvailableEnergy) 'Added by JBB
                                xEvaporationScalingFactor.Add(0) : If WeatherPointIndex.xEvaporationScalingFactor > -1 Then xEvaporationScalingFactor(Count) = WeatherRow.Value(WeatherPointIndex.xEvaporationScalingFactor) 'Added by JBB
                                DailyMinimumTemperature.Add(0) : If WeatherPointIndex.DailyMinimumTemperature > -1 Then DailyMinimumTemperature(Count) = WeatherRow.Value(WeatherPointIndex.DailyMinimumTemperature) 'Added by JBB
                                DailyMaximumTemperature.Add(0) : If WeatherPointIndex.DailyMaximumTemperature > -1 Then DailyMaximumTemperature(Count) = WeatherRow.Value(WeatherPointIndex.DailyMaximumTemperature) 'Added by JBB
                                xFetchDistanceAnemometer.Add(0) : If WeatherPointIndex.xFetchDistanceAnemometer > -1 Then xFetchDistanceAnemometer(Count) = WeatherRow.Value(WeatherPointIndex.xFetchDistanceAnemometer) 'added by JBB for wind adjust
                                xFetchDistanceCover.Add(0) : If WeatherPointIndex.xFetchDistanceCover > -1 Then xFetchDistanceCover(Count) = WeatherRow.Value(WeatherPointIndex.xFetchDistanceCover) 'added by JBB for wind adjust
                                xIBLHeightOptional.Add(0) : If WeatherPointIndex.xIBLHeightOptional > -1 Then xIBLHeightOptional(Count) = WeatherRow.Value(WeatherPointIndex.xIBLHeightOptional) 'added by JBB for wind adjust
                                xRegionalCoverHeight.Add(0) : If WeatherPointIndex.xRegionalCoverHeight > -1 Then xRegionalCoverHeight(Count) = WeatherRow.Value(WeatherPointIndex.xRegionalCoverHeight) 'added by JBB for wind adjust
                                xTypWindSpeed2m.Add(0) : If WeatherPointIndex.xTypWindSpeed2m > -1 Then xTypWindSpeed2m(Count) = WeatherRow.Value(WeatherPointIndex.xTypWindSpeed2m) 'added by JBB
                                xMonthlyETo.Add(0) : If WeatherPointIndex.xMonthlyETo > -1 Then xMonthlyETo(Count) = WeatherRow.Value(WeatherPointIndex.xMonthlyETo) 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
                                xFractionWetToday.Add(0) : If WeatherPointIndex.xFractionWetToday > -1 Then xFractionWetToday(Count) = WeatherRow.Value(WeatherPointIndex.xFractionWetToday) 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
                                Count += 1
                            End If 'added by JBB
                        End If
                    End If
                    WeatherRow = WeatherCursor.NextRow()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Loop
            System.Runtime.InteropServices.Marshal.ReleaseComObject(WeatherCursor)
        End Sub

    End Class

    Class WeatherPointIndex
        Public RecordDate As Integer
        Public ActualVaporPressure As Integer
        Public AirTemperature As Integer
        Public AnemometerReferenceHeight As Integer
        Public AirTemperatureReferenceHeight As Integer
        Public AtmosphericPressure As Integer
        Public CoverHeight As Integer
        Public CoverName As Integer
        Public Irrigation As Integer
        Public Precipitation As Integer
        Public RelativeHumidity As Integer
        Public ShortReferenceEvapotranspiration As Integer
        Public InstantaneousShortReferenceEvapotranspiration As Integer
        Public SolarRadiation As Integer
        Public WindSpeed As Integer
        Public WindFetchHeight As Integer 'Added by JBB
        Public xEvaporationScalingFactor As Integer 'added by JBB
        Public DailyAvailableEnergy As Integer 'added by JBB
        Public DailyMaximumTemperature As Integer 'added by JBB
        Public DailyMinimumTemperature As Integer 'added by JBB
        Public xFetchDistanceAnemometer As Integer 'added by JBB for wind Adjustment
        Public xFetchDistanceCover As Integer 'added by JBB for wind Adjustment
        Public xIBLHeightOptional As Integer 'added by JBB for wind adjustment
        Public xRegionalCoverHeight As Integer 'added by JBB for wind adjustment
        Public xTypWindSpeed2m As Integer ' added by JBB
        Public xMonthlyEto As Integer 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
        Public xFractionWetToday As Integer 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others


        Public Sub Initialize()
            RecordDate = -1
            ActualVaporPressure = -1
            AirTemperature = -1
            AnemometerReferenceHeight = -1
            AirTemperatureReferenceHeight = -1
            AtmosphericPressure = -1
            CoverHeight = -1
            CoverName = -1
            Irrigation = -1
            Precipitation = -1
            RelativeHumidity = -1
            ShortReferenceEvapotranspiration = -1
            InstantaneousShortReferenceEvapotranspiration = -1
            SolarRadiation = -1
            WindSpeed = -1
            WindFetchHeight = -1 'Added by JBB
            xEvaporationScalingFactor = -1 'added by JBB
            DailyAvailableEnergy = -1 ' added by JBB
            DailyMaximumTemperature = -1 'added by JBB
            DailyMinimumTemperature = -1 'added by JBB
            xFetchDistanceAnemometer = -1 'added by JBB for wind adjust
            xFetchDistanceCover = -1 'added by JBB for wind adjust
            xIBLHeightOptional = -1 'added by JBB for wind adjust
            xRegionalCoverHeight = -1 'added by JBB for wind adjust
            xTypWindSpeed2m = -1 'added by JBB
            xMonthlyEto = -1 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
            xFractionWetToday = -1 'added by JBB to adjust TEW ASCE MOP 70 ed2,2016 copying others
        End Sub
    End Class

    Class WeatherGrid
        Public Temperature As New List(Of String)
        Public SpecificHumidity As New List(Of String)
        Public Pressure As New List(Of String)
        Public WindSpeed As New List(Of String)
        Public Radiation As New List(Of String)
        Public ETDailyActual As New List(Of String)
        Public ETDailyReference As New List(Of String)
        Public ETInstantaneous As New List(Of String)
        '**********Added by JBB *****************
        Public Precipitation As New List(Of String)
        Public Irrigation As New List(Of String)
        Public Depletion As New List(Of String)
        Public ThetaV As New List(Of String)
        Public EvaporatedDepth As New List(Of String)
        Public SkinEvapDepth As New List(Of String) 'added by JBB copying others
        '***********End Added by JBB **************

        Sub Clear()
            Temperature.Clear()
            SpecificHumidity.Clear()
            Pressure.Clear()
            WindSpeed.Clear()
            Radiation.Clear()
            ETDailyActual.Clear()
            ETDailyReference.Clear()
            ETInstantaneous.Clear()
            ''**********Added by JBB *****************
            Precipitation.Clear()
            Irrigation.Clear()
            Depletion.Clear()
            ThetaV.Clear()
            EvaporatedDepth.Clear()
            SkinEvapDepth.Clear() 'added by JBB copying others
            ''***********End Added by JBB **************
        End Sub

        Function AllValues() As List(Of String)
            Dim All As New List(Of String)

            All.AddRange(Temperature)
            All.AddRange(SpecificHumidity)
            All.AddRange(Pressure)
            All.AddRange(WindSpeed)
            All.AddRange(Radiation)
            All.AddRange(ETDailyActual)
            All.AddRange(ETDailyReference)
            All.AddRange(ETInstantaneous)
            ''**********Added by JBB *****************
            All.AddRange(Precipitation)
            All.AddRange(Irrigation)
            All.AddRange(Depletion)
            All.AddRange(ThetaV)
            All.AddRange(EvaporatedDepth)
            All.AddRange(SkinEvapDepth) 'added by JBB copying others
            ''***********End Added by JBB **************
            Return All
        End Function

    End Class

    Class WeatherGridIndex
        Public Temperature As Integer
        Public SpecificHumidity As Integer
        Public Pressure As Integer
        Public WindSpeed As Integer
        Public Radiation As Integer
        Public ETDailyActual As Integer
        Public ETDailyReference As Integer
        Public ETInstantaneous As Integer
        '**********Added by JBB *****************
        Public Precipitation As Integer
        Public Irrigation As Integer
        Public Depletion As Integer
        Public ThetaV As Integer
        Public EvaporatedDepth As Integer
        Public SkinEvapDepth As Integer 'added by JBB copying others
        '***********End Added by JBB **************

        Public Sub Initialize(ByVal RecordDate As DateTime, ByVal WeatherGrid As WeatherGrid, ByVal Offset As Integer)
            Temperature = GetSameDateImageIndex(RecordDate, WeatherGrid.Temperature)
            If Temperature > -1 Then : Temperature += Offset : Offset += WeatherGrid.Temperature.Count : End If

            SpecificHumidity = GetSameDateImageIndex(RecordDate, WeatherGrid.SpecificHumidity)
            If SpecificHumidity > -1 Then : SpecificHumidity += Offset : Offset += WeatherGrid.SpecificHumidity.Count : End If

            Pressure = GetSameDateImageIndex(RecordDate, WeatherGrid.Pressure)
            If Pressure > -1 Then : Pressure += Offset : Offset += WeatherGrid.Pressure.Count : End If

            WindSpeed = GetSameDateImageIndex(RecordDate, WeatherGrid.WindSpeed)
            If WindSpeed > -1 Then : WindSpeed += Offset : Offset += WeatherGrid.WindSpeed.Count : End If

            Radiation = GetSameDateImageIndex(RecordDate, WeatherGrid.Radiation)
            If Radiation > -1 Then : Radiation += Offset : Offset += WeatherGrid.Radiation.Count : End If

            ETDailyActual = GetSameDateImageIndex(RecordDate, WeatherGrid.ETDailyActual)
            If ETDailyActual > -1 Then : ETDailyActual += Offset : Offset += WeatherGrid.ETDailyActual.Count : End If

            ETDailyReference = GetSameDateImageIndex(RecordDate, WeatherGrid.ETDailyReference)
            If ETDailyReference > -1 Then : ETDailyReference += Offset : Offset += WeatherGrid.ETDailyReference.Count : End If

            ETInstantaneous = GetSameDateImageIndex(RecordDate, WeatherGrid.ETInstantaneous)
            If ETInstantaneous > -1 Then : ETInstantaneous += Offset : Offset += WeatherGrid.ETInstantaneous.Count : End If

            '**********Added by JBB *****************
            Precipitation = GetSameDateImageIndex(RecordDate, WeatherGrid.Precipitation)
            If Precipitation > -1 Then : Precipitation += Offset : Offset += WeatherGrid.Precipitation.Count : End If

            Irrigation = GetSameDateImageIndex(RecordDate, WeatherGrid.Irrigation)
            If Irrigation > -1 Then : Irrigation += Offset : Offset += WeatherGrid.Irrigation.Count : End If

            Depletion = GetSameDateImageIndex(RecordDate, WeatherGrid.Depletion)
            If Depletion > -1 Then : Depletion += Offset : Offset += WeatherGrid.Depletion.Count : End If

            ThetaV = GetSameDateImageIndex(RecordDate, WeatherGrid.ThetaV)
            If ThetaV > -1 Then : ThetaV += Offset : Offset += WeatherGrid.ThetaV.Count : End If

            EvaporatedDepth = GetSameDateImageIndex(RecordDate, WeatherGrid.EvaporatedDepth)
            If EvaporatedDepth > -1 Then : EvaporatedDepth += Offset : Offset += WeatherGrid.EvaporatedDepth.Count : End If

            SkinEvapDepth = GetSameDateImageIndex(RecordDate, WeatherGrid.SkinEvapDepth)
            If SkinEvapDepth > -1 Then : SkinEvapDepth += Offset : Offset += WeatherGrid.SkinEvapDepth.Count : End If

            '***********End Added by JBB **************

        End Sub

    End Class

    Class Clumping_Output
        Public Clump0 As Single
        Public ClumpSun As Single
        Public ClumpView As Single
    End Class

    Class EnergyComponents_Output
        Public RnCanopy As Single
        Public RnSoil As Single
        Public RnTotal As Single
        Public HTotal As Single
        Public ETotal As Single
        Public GTotal As Single
        Public Hcanopy As Single
        Public Hsoil As Single
        Public Ecanopy As Single
        Public EVsoil As Single
        Public Stability As Integer
        Public PstTlr As Single 'Added by JBB
        Public Rcpy As Single 'added by JBB for PM
    End Class

    Class RnCoefficients_Output
        Public TauSolar As Single
        Public TauThermal As Single
        Public AlbCanopy As Single
        Public AlbSoil As Single
        Public fclear As Single
    End Class
    Class Radcoefficients
        Public DiffVis As Single
        Public DiffNIR As Single
        Public DirVis As Single
        Public DirNIR As Single
        Public fvis As Single
        Public fnir As Single
        Public fclear As Single
    End Class
    Class TInitial_Output
        Public Tsoil As Single
        Public Tcanopy As Single
        Public Tac As Single
    End Class

    Class W_Output
        Public CpRho As Single
        Public Gamma As Single
        Public Cp2 As Single
        Public Cp2P As Single
        Public W As Single
        Public Lambda As Single
        Public Rho As Single
    End Class

    Class Resistances_Output
        Public Ucanopy As Single
        Public Usoil As Single
        Public UdoZom As Single
        Public Ustar As Single
        Public Ra As Single
        Public Rex As Single
        Public Rx As Single
        Public Rah As Single
        Public Rsoil As Single
        Public y_m As Single
        Public y_h As Single
    End Class
    '********Added by JBB
    Class WaterBalance_InputDepletionOutput
        Public RecordDate As DateTime
        Public Dr_Old As Single
        Public Dr_Act As Single
        Public Dr As Single

        Sub New(ByRef RecordDate As DateTime, ByRef Dr_Old As Single, ByRef Dr_Act As Single, ByRef Dr As Single)
            Me.RecordDate = RecordDate
            Me.Dr_Old = Dr_Old
            Me.Dr_Act = Dr_Act
            Me.Dr = Dr
        End Sub

        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            SB.Append(RecordDate) '.ToString("MM/dd/yyyy  HH-mm"))'modified by JBB following others
            SB.Append("," & Dr_Old)
            SB.Append("," & Dr_Act)
            SB.Append("," & Dr)
            Return SB.ToString
        End Function

    End Class
    '*********End added by JBB

    '********Added by JBB
    Class WaterBalance_InputThetaVOutput
        Public RecordDate As DateTime
        Public ThetaVL_Old As Single
        Public ThetaVL_Act As Single
        Public ThetaVL As Single

        Sub New(ByRef RecordDate As DateTime, ByRef ThetaVL_Old As Single, ByRef ThetaVL_Act As Single, ByRef ThetaVL As Single)
            Me.RecordDate = RecordDate
            Me.ThetaVL_Old = ThetaVL_Old
            Me.ThetaVL_Act = ThetaVL_Act
            Me.ThetaVL = ThetaVL
        End Sub

        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            SB.Append(RecordDate) '.ToString("MM/dd/yyyy  HH-mm"))'modified by JBB following others
            SB.Append("," & ThetaVL_Old)
            SB.Append("," & ThetaVL_Act)
            SB.Append("," & ThetaVL)
            Return SB.ToString
        End Function

    End Class
    '*********End added by JBB

    '********Added by JBB
    Class WaterBalance_InputEvaporatedDepthOutput
        Public RecordDate As DateTime
        Public De_Old As Single
        Public De_Act As Single
        Public De As Single

        Sub New(ByRef RecordDate As DateTime, ByRef De_Old As Single, ByRef De_Act As Single, ByRef De As Single)
            Me.RecordDate = RecordDate
            Me.De_Old = De_Old
            Me.De_Act = De_Act
            Me.De = De
        End Sub

        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            SB.Append(RecordDate) '.ToString("MM/dd/yyyy  HH-mm"))'modified by JBB following others
            SB.Append("," & De_Old)
            SB.Append("," & De_Act)
            SB.Append("," & De)
            Return SB.ToString
        End Function

    End Class


    Class WaterBalance_InputSkinEvapDepthOutput 'Added by JBB copying code from above 2/9/2018
        Public RecordDate As DateTime
        Public Drew_Old As Single
        Public Drew_Act As Single
        Public Drew As Single

        Sub New(ByRef RecordDate As DateTime, ByRef Drew_Old As Single, ByRef Drew_Act As Single, ByRef Drew As Single)
            Me.RecordDate = RecordDate
            Me.Drew_Old = Drew_Old
            Me.Drew_Act = Drew_Act
            Me.Drew = Drew
        End Sub

        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            SB.Append(RecordDate) '.ToString("MM/dd/yyyy  HH-mm"))'modified by JBB following others
            SB.Append("," & Drew_Old)
            SB.Append("," & Drew_Act)
            SB.Append("," & Drew)
            Return SB.ToString
        End Function

    End Class


    Class WaterBalance_LayerOutput 'Added by JBB copying code from above 3/21/2018
        Public LyrNo As Integer
        Public Thcknss As Single
        Public FldCap As Single
        Public WPnt As Single
        Public ThtaIni As Single
        Public ThtaSat As Single
        Public Ksat As Single

        Sub New(ByRef LyrNo As Integer, ByRef Thcknss As Single, ByRef FldCap As Single, ByRef WPnt As Single, ByRef ThtaIni As Single, ByRef ThtaSat As Single, ByRef Ksat As Single)
            Me.LyrNo = LyrNo
            Me.Thcknss = Thcknss
            Me.FldCap = FldCap
            Me.WPnt = WPnt
            Me.ThtaIni = ThtaIni
            Me.ThtaSat = ThtaSat
            Me.Ksat = Ksat
        End Sub

        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            SB.Append(LyrNo) 'modified by JBB following others
            SB.Append("," & Thcknss)
            SB.Append("," & FldCap)
            SB.Append("," & WPnt)
            SB.Append("," & ThtaIni)
            SB.Append("," & ThtaSat)
            SB.Append("," & Ksat)
            Return SB.ToString
        End Function

    End Class



    '*********End added by JBB

    '********Added by JBB 2/8/2018 copying others
    Class WaterBalance_FalseKcbOutput
        Public RecordDate As DateTime
        Public CGDD As Single
        Public Kcb_Forecast As Single
        Public VI_Max As Single
        Sub New(ByRef RecordDate As DateTime, ByRef CGDD As Single, ByRef Kcb_Forecast As Single, ByRef VI_Max As Single)
            Me.RecordDate = RecordDate
            Me.CGDD = CGDD
            Me.Kcb_Forecast = Kcb_Forecast
            Me.VI_Max = VI_Max
        End Sub
        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            SB.Append(RecordDate)
            SB.Append("," & CGDD)
            SB.Append("," & Kcb_Forecast)
            SB.Append("," & VI_Max)
            Return SB.ToString
        End Function
    End Class
    '*********End added by JBB

    Class WaterBalance_ImageOverpassOutput
        '****Code below commented out by JBB
        'Public RecordDate As DateTime
        'Public Ke As single
        'Public Ks As single
        'Public Kcb_FAO As single
        'Public Kcb_VI As single
        'Public RootDepth As single
        'Public FieldCapacity As single
        'Public SoilMoisture As single
        'Public ETo As single
        'Public ET_Kcb As single
        'Public ET_EB As single
        'Public ET_As As single
        '****code below added by JBB
        Public RecordDate As DateTime
        Public Ke As Single
        Public Ks_Old As Single
        Public Ks As Single
        Public Kcb As Single
        Public Kcb_VI As Single
        Public RootDepth As Single
        Public FieldCapacity As Single
        Public WiltingPoint As Single
        Public InitialSoilMoisture As Single
        Public SoilMoisture As Single
        Public Dr_Old As Single
        Public Dr As Single
        Public ETo As Single
        Public ET_Kcb As Single
        Public ET_EB As Single
        Public ET_As As Single

        Sub New(ByRef RecordDate As DateTime, ByRef Ke As Single, ByRef Ks_Old As Single, ByRef Ks As Single, ByRef Kcb As Single, ByRef Kcb_VI As Single, ByRef RootDepth As Single, ByRef FieldCapacity As Single, ByRef WiltingPoint As Single, ByRef InitialSoilMoisture As Single, ByRef SoilMoisture As Single, ByRef Dr_Old As Single, ByRef Dr As Single, ByRef ETo As Single, ByRef ET_Kcb As Single, ByRef ET_EB As Single, ByRef ET_As As Single)
            '****Code below commented out by JBB
            'Me.RecordDate = RecordDate
            'Me.Ke = Ke
            'Me.Ks = Ks
            'Me.Kcb_FAO = Kcb_FAO
            'Me.Kcb_VI = Kcb_VI
            'Me.RootDepth = RootDepth
            'Me.FieldCapacity = FieldCapacity
            'Me.SoilMoisture = SoilMoisture
            'Me.ETo = ETo
            'Me.ET_Kcb = ET_Kcb
            'Me.ET_EB = ET_EB
            'Me.ET_As = ET_As
            '****Code below added by JBB for increased output
            Me.RecordDate = RecordDate
            Me.Ke = Ke
            Me.Ks_Old = Ks_Old
            Me.Ks = Ks
            Me.Kcb = Kcb
            Me.Kcb_VI = Kcb_VI
            Me.RootDepth = RootDepth
            Me.FieldCapacity = FieldCapacity
            Me.WiltingPoint = WiltingPoint
            Me.InitialSoilMoisture = InitialSoilMoisture
            Me.SoilMoisture = SoilMoisture
            Me.Dr_Old = Dr_Old
            Me.Dr = Dr
            Me.ETo = ETo
            Me.ET_Kcb = ET_Kcb
            Me.ET_EB = ET_EB
            Me.ET_As = ET_As
        End Sub

        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            SB.Append(RecordDate) '.ToString("MM/dd/yyyy  HH-mm"))'modified by JBB following others
            '****Code below commented out by JBB
            'SB.Append("," & Ke)
            'SB.Append("," & Ks)
            'SB.Append("," & Kcb_FAO)
            'SB.Append("," & Kcb_VI)
            'SB.Append("," & RootDepth)
            'SB.Append("," & FieldCapacity)
            'SB.Append("," & SoilMoisture)
            'SB.Append("," & ETo)
            'SB.Append("," & ET_Kcb)
            'SB.Append("," & ET_EB)
            'SB.Append("," & ET_As)
            '****Code below added by JBB for increased output
            SB.Append("," & RecordDate)
            SB.Append("," & Ke)
            SB.Append("," & Ks_Old)
            SB.Append("," & Ks)
            SB.Append("," & Kcb)
            SB.Append("," & Kcb_VI)
            SB.Append("," & RootDepth)
            SB.Append("," & FieldCapacity)
            SB.Append("," & WiltingPoint)
            SB.Append("," & InitialSoilMoisture)
            SB.Append("," & SoilMoisture)
            SB.Append("," & Dr_Old)
            SB.Append("," & Dr)
            SB.Append("," & ETo)
            SB.Append("," & ET_Kcb)
            SB.Append("," & ET_EB)
            SB.Append("," & ET_As)

            Return SB.ToString
        End Function

    End Class

    Class WaterBalance_SeasonOutput
        '*******code below commented out by JBB
        'Public RecordDate As DateTime
        'Public ETo As single
        'Public ETa As single
        'Public Depletion As single
        'Public Kcb As single
        'Public Ke As single
        'Public Ks As single
        'Public RootDepth As single
        '**********code below added by JBB for increased output following code above
        'Public RecordDate As DateTime
        'Public ETo As Single
        'Public Fc As Single
        'Public Kcb As Single
        'Public Ke As Single
        'Public Ks As Single
        'Public ETc As Single
        'Public Peff As Single
        'Public IrrNet As Single
        'Public DP As Single
        'Public CR As Single
        'Public TAW As Single
        'Public RAW As Single
        'Public TEW As Single
        'Public REW As Single
        'Public RootDepth As Single
        'Public Depletion As Single
        'Public MAD As Single
        'Public ReqIrr As Single
        'Public AWAbvMAD As Single
        'Public Zrmax As Single
        'Public θvL As Single
        'Public DPL As Single
        'Public CRL As Single
        'Public θvProf As Single
        'Public P As Single
        'Public Kr As Single
        'Public De As Single
        'Public Fw As Single
        'Public Few As Single
        'Public CGDD As Single


        Dim RecordDate As DateTime
        Dim Cover As String
        Dim YearStartWB As Integer
        Dim DoY As Integer
        Dim ETReference As Single
        Dim Precipitation As Single
        Dim Igross As Single
        Dim Tavg As Single
        Dim GDD As Single
        Dim SAVI As Single
        Dim Kcb As Single
        Dim StrSeason As String
        Dim HcPeriod As Single
        Dim U2Limited As Single
        Dim RHminLimited As Single
        Dim HcTab As Single
        Dim Fc As Single
        Dim Zri As Single
        Dim ArrDelZr1 As Single
        Dim ArrDelZr2 As Single
        Dim ArrDelZr3 As Single
        Dim ArrThtaIni0 As Single
        Dim ArrWP0 As Single
        Dim ArrFC0 As Single
        Dim VWCsat0 As Single
        Dim Ksat0 As Single
        Dim TAW As Single
        Dim Irrigation As Single
        Dim CN As Single
        Dim S_SCS As Single
        Dim InitialAbstractions As Single
        Dim ROOut As Single
        Dim EffectivePrecipitation As Single
        Dim KcMax As Single
        Dim Fw As Single
        Dim Few As Single
        Dim TEW As Single
        Dim REW As Single
        Dim Fstage1 As Single
        Dim Kt As Single
        Dim Te As Single
        Dim Tdrew As Single
        Dim Kr As Single
        Dim De As Single
        Dim DREW As Single
        Dim Ke_ns As Single
        Dim Ke_strss As Single
        Dim P As Single
        Dim RAW As Single
        Dim DelDrlastFromDelZr As Single
        Dim DrLast As Single
        Dim CR As Single
        Dim ArrThta0 As Single
        Dim ArrTau0 As Single
        Dim DP As Single
        Dim Fexc As Single
        Dim DrLimit As Single
        Dim Dr As Single
        Dim Ks As Single
        Dim Kc As Single
        Dim ETc As Single
        Dim ETcAdjOut As Single
        Dim EvapNS As Single
        Dim Evap As Single
        Dim Tc_ns As Single
        Dim Tc_strss As Single
        Dim ArrDelZL1 As Single
        Dim ArrDelZL2 As Single
        Dim ArrDelZL3 As Single
        Dim ZL As Single
        Dim TAWLower As Single
        Dim WPrPrev As Single
        Dim FCrPrev As Single
        Dim ArrThtaPrev1 As Single
        Dim ArrThtaPrev2 As Single
        Dim ArrThtaPrev3 As Single
        Dim DPtot As Single
        Dim θvL As Single
        Dim ArrThta1 As Single
        Dim ArrThta2 As Single
        Dim ArrThta3 As Single
        Dim ProfileAvgθv As Single
        Dim AssimExcess As Single
        Dim AssimExcessL As Single
        Dim WBClosureFull As Single
        Dim WBClosureRZ As Single
        Dim MAD As Single
        Dim RequiredGrossIrrigation As Single
        Dim AvailWaterAboveMAD As Single
        Dim Srel As Single
        Dim Kst As Single
        Dim SumKcOut As Single
        Dim SumETrOut As Single
        Dim SumPcpOut As Single
        Dim SumROOut As Single
        Dim SumInetOut As Single
        Dim SumETcAdjOut As Single
        Dim SumDPOut As Single
        Dim SumDPLOut As Single
        Dim SumFexcOut As Single
        Dim SumAddFromDelZr As Single
        Dim SumAssimExcess As Single
        Dim SumAssimExcessL As Single
        Dim CumWBClosureFull As Single
        Dim CumWBClosureRZ As Single







        '*******code below commented out by JBB
        'Sub New(ByRef RecordDate As DateTime, ByRef ETo As single, ByRef ETa As single, ByRef Depletion As single, ByRef Kcb As single, ByRef Ke As single, ByRef Ks As single, ByRef RootDepth As single)
        '    Me.RecordDate = RecordDate
        '    Me.ETo = ETo
        '    Me.ETa = ETa
        '    Me.Depletion = Depletion
        '    Me.Kcb = Kcb
        '    Me.Ke = Ke
        '    Me.Ks = Ks
        '    Me.RootDepth = RootDepth
        'End Sub
        '**********code below added by JBB for increased output following code above
        ' Sub New(ByRef RecordDate As DateTime, ByRef ETo As Single, ByRef Fc As Single, ByRef Kcb As Single, ByRef Ke As Single, ByRef Ks As Single, ByRef ETc As Single, ByRef Peff As Single, ByRef IrrNet As Single, ByRef DP As Single, ByRef CR As Single, ByRef TAW As Single, ByRef RAW As Single, ByRef TEW As Single, ByRef REW As Single, ByRef RootDepth As Single, ByRef Depletion As Single, ByRef MAD As Single, ByRef ReqIrr As Single, ByRef AWAbvMAD As Single, ByRef Zrmax As Single, ByRef θvL As Single, ByRef DPL As Single, ByRef CRL As Single, ByRef θvProf As Single, ByRef P As Single, ByRef Kr As Single, ByRef De As Single, ByRef Fw As Single, ByRef Few As Single, ByRef CGDD As Single)
        Sub New(ByRef RecordDate As DateTime, ByRef Cover As String, ByRef YearStartWB As Integer, ByRef DoY As Integer, ByRef ETReference As Single, ByRef Precipitation As Single, ByRef Igross As Single, ByRef Tavg As Single, ByRef GDD As Single, ByRef SAVI As Single, ByRef Kcb As Single, ByRef StrSeason As String, ByRef HcPeriod As Single, ByRef U2Limited As Single, ByRef RHminLimited As Single, ByRef HcTab As Single, ByRef Fc As Single, ByRef Zri As Single, ByRef ArrDelZr1 As Single, ByRef ArrDelZr2 As Single, ByRef ArrDelZr3 As Single, ByRef ArrThtaIni0 As Single, ByRef ArrWP0 As Single, ByRef ArrFC0 As Single, ByRef VWCsat0 As Single, ByRef Ksat0 As Single, ByRef TAW As Single, ByRef Irrigation As Single, ByRef CN As Single, ByRef S_SCS As Single, ByRef InitialAbstractions As Single, ByRef ROOut As Single, ByRef EffectivePrecipitation As Single, ByRef KcMax As Single, ByRef Fw As Single, ByRef Few As Single, ByRef TEW As Single, ByRef REW As Single, ByRef Fstage1 As Single, ByRef Kt As Single, ByRef Te As Single, ByRef Tdrew As Single, ByRef Kr As Single, ByRef De As Single, ByRef DREW As Single, ByRef Ke_ns As Single, ByRef Ke_strss As Single, ByRef P As Single, ByRef RAW As Single, ByRef DelDrlastFromDelZr As Single, ByRef DrLast As Single, ByRef CR As Single, ByRef ArrThta0 As Single, ByRef ArrTau0 As Single, ByRef DP As Single, ByRef Fexc As Single, ByRef DrLimit As Single, ByRef Dr As Single, ByRef Ks As Single, ByRef Kc As Single, ByRef ETc As Single, ByRef ETcAdjOut As Single, ByRef EvapNS As Single, ByRef Evap As Single, ByRef Tc_ns As Single, ByRef Tc_strss As Single, ByRef ArrDelZL1 As Single, ByRef ArrDelZL2 As Single, ByRef ArrDelZL3 As Single, ByRef ZL As Single, ByRef TAWLower As Single, ByRef WPrPrev As Single, ByRef FCrPrev As Single, ByRef ArrThtaPrev1 As Single, ByRef ArrThtaPrev2 As Single, ByRef ArrThtaPrev3 As Single, ByRef DPtot As Single, ByRef θvL As Single, ByRef ArrThta1 As Single, ByRef ArrThta2 As Single, ByRef ArrThta3 As Single, ByRef ProfileAvgθv As Single, ByRef AssimExcess As Single, ByRef AssimExcessL As Single, ByRef WBClosureFull As Single, ByRef WBClosureRZ As Single, ByRef MAD As Single, ByRef RequiredGrossIrrigation As Single, ByRef AvailWaterAboveMAD As Single, ByRef Srel As Single, ByRef Kst As Single, ByRef SumKcOut As Single, ByRef SumETrOut As Single, ByRef SumPcpOut As Single, ByRef SumROOut As Single, ByRef SumInetOut As Single, ByRef SumETcAdjOut As Single, ByRef SumDPOut As Single, ByRef SumDPLOut As Single, ByRef SumFexcOut As Single, ByRef SumAddFromDelZr As Single, ByRef SumAssimExcess As Single, ByRef SumAssimExcessL As Single, ByRef CumWBClosureFull As Single, ByRef CumWBClosureRZ As Single)
            'Me.RecordDate = RecordDate
            'Me.ETo = ETo
            'Me.Fc = Fc
            'Me.Kcb = Kcb
            'Me.Ke = Ke
            'Me.Ks = Ks
            'Me.ETc = ETc
            'Me.Peff = Peff
            'Me.IrrNet = IrrNet
            'Me.DP = DP
            'Me.CR = CR
            'Me.TAW = TAW
            'Me.RAW = RAW
            'Me.TEW = TEW
            'Me.REW = REW
            'Me.RootDepth = RootDepth
            'Me.Depletion = Depletion
            'Me.MAD = MAD
            'Me.ReqIrr = ReqIrr
            'Me.AWAbvMAD = AWAbvMAD
            'Me.Zrmax = Zrmax
            'Me.θvL = θvL
            'Me.DPL = DPL
            'Me.CRL = CRL
            'Me.θvProf = θvProf
            'Me.P = P
            'Me.Kr = Kr
            'Me.De = De
            'Me.Fw = Fw
            'Me.Few = Few
            'Me.CGDD = CGDD


            Me.RecordDate = RecordDate
            Me.Cover = Cover
            Me.YearStartWB = YearStartWB
            Me.DoY = DoY
            Me.ETReference = ETReference
            Me.Precipitation = Precipitation
            Me.Igross = Igross
            Me.Tavg = Tavg
            Me.GDD = GDD
            Me.SAVI = SAVI
            Me.Kcb = Kcb
            Me.StrSeason = StrSeason
            Me.HcPeriod = HcPeriod
            Me.U2Limited = U2Limited
            Me.RHminLimited = RHminLimited
            Me.HcTab = HcTab
            Me.Fc = Fc
            Me.Zri = Zri
            Me.ArrDelZr1 = ArrDelZr1
            Me.ArrDelZr2 = ArrDelZr2
            Me.ArrDelZr3 = ArrDelZr3
            Me.ArrThtaIni0 = ArrThtaIni0
            Me.ArrWP0 = ArrWP0
            Me.ArrFC0 = ArrFC0
            Me.VWCsat0 = VWCsat0
            Me.Ksat0 = Ksat0
            Me.TAW = TAW
            Me.Irrigation = Irrigation
            Me.CN = CN
            Me.S_SCS = S_SCS
            Me.InitialAbstractions = InitialAbstractions
            Me.ROOut = ROOut
            Me.EffectivePrecipitation = EffectivePrecipitation
            Me.KcMax = KcMax
            Me.Fw = Fw
            Me.Few = Few
            Me.TEW = TEW
            Me.REW = REW
            Me.Fstage1 = Fstage1
            Me.Kt = Kt
            Me.Te = Te
            Me.Tdrew = Tdrew
            Me.Kr = Kr
            Me.De = De
            Me.DREW = DREW
            Me.Ke_ns = Ke_ns
            Me.Ke_strss = Ke_strss
            Me.P = P
            Me.RAW = RAW
            Me.DelDrlastFromDelZr = DelDrlastFromDelZr
            Me.DrLast = DrLast
            Me.CR = CR
            Me.ArrThta0 = ArrThta0
            Me.ArrTau0 = ArrTau0
            Me.DP = DP
            Me.Fexc = Fexc
            Me.DrLimit = DrLimit
            Me.Dr = Dr
            Me.Ks = Ks
            Me.Kc = Kc
            Me.ETc = ETc
            Me.ETcAdjOut = ETcAdjOut
            Me.EvapNS = EvapNS
            Me.Evap = Evap
            Me.Tc_ns = Tc_ns
            Me.Tc_strss = Tc_strss
            Me.ArrDelZL1 = ArrDelZL1
            Me.ArrDelZL2 = ArrDelZL2
            Me.ArrDelZL3 = ArrDelZL3
            Me.ZL = ZL
            Me.TAWLower = TAWLower
            Me.WPrPrev = WPrPrev
            Me.FCrPrev = FCrPrev
            Me.ArrThtaPrev1 = ArrThtaPrev1
            Me.ArrThtaPrev2 = ArrThtaPrev2
            Me.ArrThtaPrev3 = ArrThtaPrev3
            Me.DPtot = DPtot
            Me.θvL = θvL
            Me.ArrThta1 = ArrThta1
            Me.ArrThta2 = ArrThta2
            Me.ArrThta3 = ArrThta3
            Me.ProfileAvgθv = ProfileAvgθv
            Me.AssimExcess = AssimExcess
            Me.AssimExcessL = AssimExcessL
            Me.WBClosureFull = WBClosureFull
            Me.WBClosureRZ = WBClosureRZ
            Me.MAD = MAD
            Me.RequiredGrossIrrigation = RequiredGrossIrrigation
            Me.AvailWaterAboveMAD = AvailWaterAboveMAD
            Me.Srel = Srel
            Me.Kst = Kst
            Me.SumKcOut = SumKcOut
            Me.SumETrOut = SumETrOut
            Me.SumPcpOut = SumPcpOut
            Me.SumROOut = SumROOut
            Me.SumInetOut = SumInetOut
            Me.SumETcAdjOut = SumETcAdjOut
            Me.SumDPOut = SumDPOut
            Me.SumDPLOut = SumDPLOut
            Me.SumFexcOut = SumFexcOut
            Me.SumAddFromDelZr = SumAddFromDelZr
            Me.SumAssimExcess = SumAssimExcess
            Me.SumAssimExcessL = SumAssimExcessL
            Me.CumWBClosureFull = CumWBClosureFull
            Me.CumWBClosureRZ = CumWBClosureRZ

        End Sub

        Public Function WriteDelimited() As String
            Dim SB As New System.Text.StringBuilder
            'SB.Append(RecordDate.ToString("MM/dd/yyyy  HH-mm"))'Commented out by JBB 2/8/2018      
            SB.Append(RecordDate) 'added by JBB 2/8/2018 copying the other lines
            '*****Code below commented by by JBB
            'SB.Append("," & ETo)
            'SB.Append("," & ETc)
            'SB.Append("," & Depletion)
            'SB.Append("," & Kcb)
            'SB.Append("," & Ke)
            'SB.Append("," & Ks)
            'SB.Append("," & RootDepth)

            '****Code below added by JBB for increased output
            'SB.Append("," & ETo)
            'SB.Append("," & Fc)
            'SB.Append("," & Kcb)
            'SB.Append("," & Ke)
            'SB.Append("," & Ks)
            'SB.Append("," & ETc)
            'SB.Append("," & Peff)
            'SB.Append("," & IrrNet)
            'SB.Append("," & DP)
            'SB.Append("," & CR)
            'SB.Append("," & TAW)
            'SB.Append("," & RAW)
            'SB.Append("," & TEW)
            'SB.Append("," & REW)
            'SB.Append("," & RootDepth)
            'SB.Append("," & Depletion)
            'SB.Append("," & MAD)
            'SB.Append("," & ReqIrr)
            'SB.Append("," & AWAbvMAD)
            'SB.Append("," & Zrmax)
            'SB.Append("," & θvL)
            'SB.Append("," & DPL)
            'SB.Append("," & CRL)
            'SB.Append("," & θvProf)
            'SB.Append("," & P)
            'SB.Append("," & Kr)
            'SB.Append("," & De)
            'SB.Append("," & Fw)
            'SB.Append("," & Few)
            'SB.Append("," & CGDD)


            SB.Append("," & Cover)
            SB.Append("," & YearStartWB)
            SB.Append("," & DoY)
            SB.Append("," & ETReference)
            SB.Append("," & Precipitation)
            SB.Append("," & Igross)
            SB.Append("," & Tavg)
            SB.Append("," & GDD)
            SB.Append("," & SAVI)
            SB.Append("," & Kcb)
            SB.Append("," & StrSeason)
            SB.Append("," & HcPeriod)
            SB.Append("," & U2Limited)
            SB.Append("," & RHminLimited)
            SB.Append("," & HcTab)
            SB.Append("," & Fc)
            SB.Append("," & Zri)
            SB.Append("," & ArrDelZr1)
            SB.Append("," & ArrDelZr2)
            SB.Append("," & ArrDelZr3)
            SB.Append("," & ArrThtaIni0)
            SB.Append("," & ArrWP0)
            SB.Append("," & ArrFC0)
            SB.Append("," & VWCsat0)
            SB.Append("," & Ksat0)
            SB.Append("," & TAW)
            SB.Append("," & Irrigation)
            SB.Append("," & CN)
            SB.Append("," & S_SCS)
            SB.Append("," & InitialAbstractions)
            SB.Append("," & ROOut)
            SB.Append("," & EffectivePrecipitation)
            SB.Append("," & KcMax)
            SB.Append("," & Fw)
            SB.Append("," & Few)
            SB.Append("," & TEW)
            SB.Append("," & REW)
            SB.Append("," & Fstage1)
            SB.Append("," & Kt)
            SB.Append("," & Te)
            SB.Append("," & Tdrew)
            SB.Append("," & Kr)
            SB.Append("," & De)
            SB.Append("," & DREW)
            SB.Append("," & Ke_ns)
            SB.Append("," & Ke_strss)
            SB.Append("," & P)
            SB.Append("," & RAW)
            SB.Append("," & DelDrlastFromDelZr)
            SB.Append("," & DrLast)
            SB.Append("," & CR)
            SB.Append("," & ArrThta0)
            SB.Append("," & ArrTau0)
            SB.Append("," & DP)
            SB.Append("," & Fexc)
            SB.Append("," & DrLimit)
            SB.Append("," & Dr)
            SB.Append("," & Ks)
            SB.Append("," & Kc)
            SB.Append("," & ETc)
            SB.Append("," & ETcAdjOut)
            SB.Append("," & EvapNS)
            SB.Append("," & Evap)
            SB.Append("," & Tc_ns)
            SB.Append("," & Tc_strss)
            SB.Append("," & ArrDelZL1)
            SB.Append("," & ArrDelZL2)
            SB.Append("," & ArrDelZL3)
            SB.Append("," & ZL)
            SB.Append("," & TAWLower)
            SB.Append("," & WPrPrev)
            SB.Append("," & FCrPrev)
            SB.Append("," & ArrThtaPrev1)
            SB.Append("," & ArrThtaPrev2)
            SB.Append("," & ArrThtaPrev3)
            SB.Append("," & DPtot)
            SB.Append("," & θvL)
            SB.Append("," & ArrThta1)
            SB.Append("," & ArrThta2)
            SB.Append("," & ArrThta3)
            SB.Append("," & ProfileAvgθv)
            SB.Append("," & AssimExcess)
            SB.Append("," & AssimExcessL)
            SB.Append("," & WBClosureFull)
            SB.Append("," & WBClosureRZ)
            SB.Append("," & MAD)
            SB.Append("," & RequiredGrossIrrigation)
            SB.Append("," & AvailWaterAboveMAD)
            SB.Append("," & Srel)
            SB.Append("," & Kst)
            SB.Append("," & SumKcOut)
            SB.Append("," & SumETrOut)
            SB.Append("," & SumPcpOut)
            SB.Append("," & SumROOut)
            SB.Append("," & SumInetOut)
            SB.Append("," & SumETcAdjOut)
            SB.Append("," & SumDPOut)
            SB.Append("," & SumDPLOut)
            SB.Append("," & SumFexcOut)
            SB.Append("," & SumAddFromDelZr)
            SB.Append("," & SumAssimExcess)
            SB.Append("," & SumAssimExcessL)
            SB.Append("," & CumWBClosureFull)
            SB.Append("," & CumWBClosureRZ)

            Return SB.ToString
        End Function

    End Class

    Class Bioproperties
        Public AlphaVIS As Single
        Public AlphaNIR As Single
        Public AlphaTIR As Single
        Public AlphaLeafVIS As Single
        Public AlphaLeafNIR As Single
        Public AlphaLeafTIR As Single
        Public AlphaDeadVIS As Single
        Public AlphaDeadNIR As Single
        Public AlphaDeadTIR As Single
        'Public AlphaLongwave As single
        Public s As Single
        Public Wc As Single
        Public fg As Single
        Public HcMin As Single
        Public HcMax As Single
        Public LAITable As Single
        'Added by JBB
        Public EmissSoilVIS As Single
        Public EmissSoilNIR As Single
        Public EmissSoilTIR As Single
        Public SoilHtFlxAg As Single

        Public ClumpD As Single
        Public aPTInput As Single
        Public RcIniInput As Single
        Public RcMaxInput As Single
        'Public MaxIterInput As Single
        Public FcVIInput As Single
        Public MinVIInput As Single
        Public MaxVIInput As Single
        Public ExpVIInput As Single
        'End added by JBB
    End Class

#End Region

End Module
