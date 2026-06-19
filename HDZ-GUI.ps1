#requires -Version 7.0
# =========================================================================
# HDZ STUDIO — interfaz adaptativa para HDZnew.ps1 (WPF, tema oscuro)
# =========================================================================
# Al seleccionar la carpeta, ANALIZA los vídeos con ffprobe en segundo plano
# (códec, resolución, HDR/Dolby Vision, parejas híbridas, DTS, pistas sin
# idioma, PGS, idiomas de subtítulos, ambigüedades) y adapta cada tarjeta:
#  - desactiva las opciones sin efecto en ese lote (con el motivo),
#  - anota cada opción con lo detectado (contadores, archivos),
#  - sugiere título / año / serie a partir de los nombres de archivo,
#  - muestra los idiomas de subtítulos REALES para el filtro personalizado.
#
# La GUI decide TODAS las opciones: el montaje corre sin preguntar nada por
# consola. Ejecútala con doble clic en "HDZ Studio.cmd".
# =========================================================================

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # solo para el diálogo de carpeta

# Herramientas portables incluidas: si existe la carpeta bin\ junto al script, la anteponemos
# al PATH para que el programa use SIEMPRE esas versiones (probadas) en vez de las del sistema
# (resuelve "no está en el PATH" y "versiones antiguas"). Si no hay bin\, se usa el PATH normal.
$rutaBin = Join-Path $PSScriptRoot "bin"
if ((Test-Path -LiteralPath $rutaBin) -and ($env:PATH -notlike "*$rutaBin*")) { $env:PATH = "$rutaBin;$env:PATH" }

# Versión instalada (marcador en VERSION.txt; respaldo si no existe) y memoria de la versión
# que el usuario ya descartó actualizar (para no volver a avisarle de la MISMA).
$script:HDZVersion = "1.0.0"
$script:ajusteVerDescartada = ""
$script:bienvenidaVista = $false   # ¿ya se mostró la ventana de bienvenida (primer arranque)?

$rutaScript  = Join-Path $PSScriptRoot "HDZnew.ps1"
$rutaTorrentMod = Join-Path $PSScriptRoot "HDZ-Torrent.ps1"
$rutaAjustes = Join-Path $env:APPDATA "HDZ-GUI.settings.json"
$script:modoTest = ($env:HDZ_GUI_TEST -eq "1")

# Módulo compartido de creación de torrents (lo usa también HDZnew.ps1)
if (Test-Path -LiteralPath $rutaTorrentMod) { . $rutaTorrentMod }

# =========================================================================
# CATÁLOGOS Y NORMALIZACIÓN DE IDIOMAS (espejo de Get-LanguageCode del script)
# =========================================================================
$idiomasFiltro = @(
    @{Cod="es";     Nom="Castellano"},
    @{Cod="es-419"; Nom="Latino"},
    @{Cod="eng";    Nom="Inglés"},
    @{Cod="cat";    Nom="Catalán"},
    @{Cod="glg";    Nom="Gallego"},
    @{Cod="eus";    Nom="Euskera"},
    @{Cod="fre";    Nom="Francés"},
    @{Cod="ger";    Nom="Alemán"},
    @{Cod="ita";    Nom="Italiano"},
    @{Cod="por";    Nom="Portugués"},
    @{Cod="jpn";    Nom="Japonés"},
    @{Cod="chi";    Nom="Chino"},
    @{Cod="kor";    Nom="Coreano"}
)
$nombresIdioma = @{
    "es"="Castellano"; "es-419"="Latino"; "eng"="Inglés"; "fre"="Francés"; "ger"="Alemán"
    "ita"="Italiano"; "por"="Portugués"; "jpn"="Japonés"; "chi"="Chino"; "kor"="Coreano"
    "cat"="Catalán"; "glg"="Gallego"; "eus"="Euskera"; "und"="Indeterminado"
}
function NombreIdioma($cod) { if ($nombresIdioma.ContainsKey("$cod")) { $nombresIdioma["$cod"] } else { "$cod" } }

function ConvCanon($lang, $title = "") {
    $l = "$lang".ToLower().Trim()
    if ([string]::IsNullOrWhiteSpace($l) -or $l -eq "und") { return "und" }
    if ($l -match "^es-(419|mx|ar|co|cl|pe|ve|la|us)$" -or $l -eq "lat") { return "es-419" }
    if ($l -in @("es", "spa", "cas") -or $l -match "^es-") {
        if ("$title" -match "(?i)latino|latinoam|latam|hispanoam") { return "es-419" }
        return "es"
    }
    if ($l -in @("en", "eng") -or $l -match "^en-") { return "eng" }
    if ($l -in @("fr", "fra", "fre") -or $l -match "^fr-") { return "fre" }
    if ($l -in @("de", "ger", "deu") -or $l -match "^de-") { return "ger" }
    if ($l -in @("it", "ita")) { return "ita" }
    if ($l -in @("pt", "por") -or $l -match "^pt-") { return "por" }
    if ($l -in @("ja", "jpn", "jp")) { return "jpn" }
    if ($l -in @("zh", "chi", "zho", "cmn") -or $l -match "^zh-") { return "chi" }
    if ($l -in @("ko", "kor")) { return "kor" }
    if ($l -in @("ca", "cat")) { return "cat" }
    if ($l -in @("gl", "glg")) { return "glg" }
    if ($l -in @("eu", "eus", "baq")) { return "eus" }
    return $l
}

$opcionesIdiomaUnd = @(@{T="Dejar como indeterminado (und)"; V="und"}) +
                     @($idiomasFiltro | ForEach-Object { @{T="$($_.Nom)  ·  $($_.Cod)"; V=$_.Cod} })

$plataformasGui = @("NF — Netflix","AMZN — Amazon Prime","DSNP — Disney+","MVSTP — Movistar+",
    "SKST — SkyShowtime","FLMN — Filmin","ATVP — Apple TV+","iT — iTunes","HMAX — HBO Max",
    "RTVP — RTVE Play","RKTN — Rakuten")
$formatosGui = @("UHDFull","BDFull","UHDRemux","BDRemux","UHDRip","BDRip","MHD","Remastered")

# Mismo normalizador de agrupación que usa HDZnew.ps1 para detectar parejas híbridas
$normalizadorAgrupacion = '(?i)[\._\s\-\(\)]+(Dolby[\._\s\-]Vision|DolbyVision|HDR10\+|HDR10|HDR|DV|2160p|1080p|UHD|WEB[\._\s\-]?DL|Apple[\._\s\-]?TV|iTunes|Disney\+?|Atmos|TrueHD|DTS[\._\s\-]?HD[\._\s\-]?MA|DTS[\._\s\-]?HD|DTS[\._\s\-]?X|DTS|E[\._\s\-]?AC[\._\s\-]?3|EAC3|AC3|DDP\+?|DD\+|DDP|DD|AAC|FLAC|MA|HRA)(?=[\._\s\-\(\)]|$)'

# =========================================================================
# XAML
# =========================================================================
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="HDZ Studio" Width="1100" Height="760" MinWidth="960" MinHeight="640"
        WindowStartupLocation="CenterScreen" Background="#0C0C0F"
        FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True"
        SnapsToDevicePixels="True" AllowDrop="True"
        TextOptions.TextFormattingMode="Display" TextOptions.TextRenderingMode="ClearType">
  <Window.Resources>

    <SolidColorBrush x:Key="BgBrush"          Color="#0C0C0F"/>
    <SolidColorBrush x:Key="SidebarBrush"     Color="#09090B"/>
    <SolidColorBrush x:Key="CardBrush"        Color="#131316"/>
    <SolidColorBrush x:Key="CardBorderBrush"  Color="#242429"/>
    <SolidColorBrush x:Key="FieldBrush"       Color="#0E0E11"/>
    <SolidColorBrush x:Key="FieldBorderBrush" Color="#2E2E35"/>
    <SolidColorBrush x:Key="ChipBrush"        Color="#1A1A1F"/>
    <SolidColorBrush x:Key="ChipBorderBrush"  Color="#2E2E35"/>
    <SolidColorBrush x:Key="TextBrush"        Color="#EDEDEF"/>
    <SolidColorBrush x:Key="SubBrush"         Color="#9C9CA8"/>
    <SolidColorBrush x:Key="MutedBrush"       Color="#5E5E66"/>
    <SolidColorBrush x:Key="AccentBrush"      Color="#D92B2B"/>
    <SolidColorBrush x:Key="AccentHoverBrush" Color="#EF4444"/>
    <SolidColorBrush x:Key="OkBrush"          Color="#86D97C"/>
    <SolidColorBrush x:Key="WarnBrush"        Color="#F2C94C"/>
    <SolidColorBrush x:Key="ErrBrush"         Color="#FF6B81"/>

    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="9"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid Background="Transparent">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.DecreaseRepeatButton>
                  <RepeatButton Command="ScrollBar.PageUpCommand">
                    <RepeatButton.Template><ControlTemplate TargetType="RepeatButton"><Rectangle Fill="Transparent"/></ControlTemplate></RepeatButton.Template>
                  </RepeatButton>
                </Track.DecreaseRepeatButton>
                <Track.IncreaseRepeatButton>
                  <RepeatButton Command="ScrollBar.PageDownCommand">
                    <RepeatButton.Template><ControlTemplate TargetType="RepeatButton"><Rectangle Fill="Transparent"/></ControlTemplate></RepeatButton.Template>
                  </RepeatButton>
                </Track.IncreaseRepeatButton>
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template><ControlTemplate TargetType="Thumb"><Border Background="#3A3A42" CornerRadius="4" Margin="2,1"/></ControlTemplate></Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource CardBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource CardBorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Padding" Value="18,15"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>

    <Style x:Key="H" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="D" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource SubBrush}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,3,0,0"/>
    </Style>
    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource SubBrush}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Margin" Value="0,12,0,5"/>
    </Style>
    <Style x:Key="Bdg" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="TextAlignment" Value="Right"/>
      <Setter Property="MaxWidth" Value="340"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="12,0,0,0"/>
    </Style>

    <Style x:Key="Chip" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="bd" Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                    BorderThickness="1" CornerRadius="16" Padding="14,6.5">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentBrush}"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ChipToggle" TargetType="ToggleButton">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToggleButton">
            <Border x:Name="bd" Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                    BorderThickness="1" CornerRadius="16" Padding="14,6.5">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentBrush}"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Switch" TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Grid Width="42" Height="23" VerticalAlignment="Center">
                <Border x:Name="track" CornerRadius="11.5" Background="#2E2E35"/>
                <Ellipse x:Name="thumb" Width="15" Height="15" Fill="#9C9CA8" HorizontalAlignment="Left" Margin="4,0,0,0"/>
              </Grid>
              <ContentPresenter Margin="11,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="track" Property="Background" Value="{StaticResource AccentBrush}"/>
                <Setter TargetName="thumb" Property="Fill" Value="White"/>
                <Setter TargetName="thumb" Property="HorizontalAlignment" Value="Right"/>
                <Setter TargetName="thumb" Property="Margin" Value="0,0,4,0"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="track" Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Input" TargetType="TextBox">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="CaretBrush" Value="White"/>
      <Setter Property="Background" Value="{StaticResource FieldBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource FieldBorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#56565E"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ComboItem" TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" Background="Transparent" CornerRadius="6" Margin="4,1" Padding="10,7">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#2A2A30"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentBrush}"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Combo" TargetType="ComboBox">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="ItemContainerStyle" Value="{StaticResource ComboItem}"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <Border x:Name="bd" Background="{StaticResource FieldBrush}" BorderBrush="{StaticResource FieldBorderBrush}"
                      BorderThickness="1" CornerRadius="8"/>
              <ToggleButton ClickMode="Press" Focusable="False" Background="Transparent"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton"><Border Background="Transparent"/></ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter Margin="12,7,32,7" VerticalAlignment="Center" IsHitTestVisible="False"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                TextElement.Foreground="{StaticResource TextBrush}"/>
              <Path Data="M0,0 L4.5,4.5 L9,0" Stroke="#9C9CA8" StrokeThickness="1.6" StrokeLineJoin="Round"
                    HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,1,13,0" IsHitTestVisible="False"/>
              <Popup x:Name="PART_Popup" IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom"
                     AllowsTransparency="True" PopupAnimation="Slide" Focusable="False">
                <Border Background="#18181C" BorderBrush="#33333A" BorderThickness="1" CornerRadius="8"
                        Margin="0,5,0,8" Padding="0,4" MaxHeight="300"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <Border.Effect><DropShadowEffect BlurRadius="18" ShadowDepth="3" Opacity="0.45" Color="#000000"/></Border.Effect>
                  <ScrollViewer VerticalScrollBarVisibility="Auto"><ItemsPresenter/></ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#56565E"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Nav" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#A8A8B2"/>
      <Setter Property="FontSize" Value="13.5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Padding" Value="13,10"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="bd" CornerRadius="9" Background="Transparent" Padding="{TemplateBinding Padding}" Margin="0,2">
              <Grid>
                <Rectangle x:Name="ind" Width="3" RadiusX="1.5" RadiusY="1.5" Fill="{StaticResource AccentBrush}"
                           HorizontalAlignment="Left" Margin="-7,2,0,2" Visibility="Collapsed"/>
                <ContentPresenter VerticalAlignment="Center"
                                  HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#141417"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#1E1E22"/>
                <Setter TargetName="ind" Property="Visibility" Value="Visible"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Primary" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{StaticResource AccentBrush}" CornerRadius="10" Padding="26,12">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#A81E1E"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <!-- Variante compacta del botón rojo (mismo color, menos relleno) -->
    <Style x:Key="PrimarySmall" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{StaticResource AccentBrush}" CornerRadius="8" Padding="15,8">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#A81E1E"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Ghost" TargetType="Button">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                    BorderThickness="1" CornerRadius="8" Padding="16,8">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="GhostMini" TargetType="Button">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                    BorderThickness="1" CornerRadius="6" Padding="11,4">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Botón-icono cuadrado para la barra BBCode -->
    <Style x:Key="IconBtn" TargetType="Button">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Width" Value="32"/>
      <Setter Property="Height" Value="30"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                    BorderThickness="1" CornerRadius="6">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentBrush}"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Casilla de selección de archivo (lista de análisis) -->
    <Style x:Key="Check" TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Border x:Name="box" Width="17" Height="17" CornerRadius="4" VerticalAlignment="Center"
                      Background="{StaticResource FieldBrush}" BorderBrush="{StaticResource FieldBorderBrush}" BorderThickness="1">
                <Path x:Name="tick" Data="M1,5 L4,8 L9,1" Stroke="White" StrokeThickness="2"
                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                      Stretch="Uniform" Margin="3" Visibility="Collapsed"/>
              </Border>
              <ContentPresenter Margin="9,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="box" Property="Background" Value="{StaticResource AccentBrush}"/>
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                <Setter TargetName="tick" Property="Visibility" Value="Visible"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Asa de redimensión (esquina inferior derecha) para los campos de texto grandes -->
    <Style x:Key="ResizeGrip" TargetType="Thumb">
      <Setter Property="Width" Value="18"/>
      <Setter Property="Height" Value="18"/>
      <Setter Property="HorizontalAlignment" Value="Right"/>
      <Setter Property="VerticalAlignment" Value="Bottom"/>
      <Setter Property="Cursor" Value="SizeNS"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border Background="Transparent" ToolTip="Arrastra para cambiar el alto">
              <Path Stroke="#7A7A86" StrokeThickness="1.4" Margin="0,0,3,3"
                    HorizontalAlignment="Right" VerticalAlignment="Bottom"
                    Data="M 13,5 L 5,13 M 13,9 L 9,13 M 13,13 L 12.5,13.5"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid Background="#0C0C0F">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <!-- ============ SIDEBAR ============ -->
    <Border x:Name="sidebar" Width="216" Background="{StaticResource SidebarBrush}">
      <DockPanel Margin="12,18,12,18">
        <StackPanel DockPanel.Dock="Top">
          <!-- El logo ya incluye «STUDIO», por eso no hay texto «Studio» aparte debajo.
               BitmapScalingMode=HighQuality (Fant): el archivo viene a 464px y el control lo muestra
               a ~180px (reducción suave ~2.6×). El escalador por defecto de WPF NO promedia y deja
               grano en los bordes curvos (R, O); Fant promedia y los deja limpios sin emborronar. -->
          <Image x:Name="imgLogo" MaxHeight="68" HorizontalAlignment="Left" Margin="6,0,6,20"
                 Stretch="Uniform" Visibility="Collapsed"
                 RenderOptions.BitmapScalingMode="HighQuality"/>
          <StackPanel x:Name="panLogoTexto" Orientation="Horizontal" Margin="8,0,0,20">
            <TextBlock Text="HD" FontSize="27" FontWeight="Bold" Foreground="{StaticResource AccentBrush}"/>
            <TextBlock Text="ZERO" FontSize="27" FontWeight="Bold" Foreground="{StaticResource TextBrush}" Margin="2,0,0,0"/>
          </StackPanel>

          <!-- RAMA 1: creación del archivo (montaje) -->
          <TextBlock x:Name="lblSecMontaje" Text="① CREAR EL ARCHIVO" Foreground="{StaticResource AccentBrush}" FontSize="10.5"
                     FontWeight="Bold" Margin="10,0,0,7"/>
          <RadioButton x:Name="navGeneral"  Style="{StaticResource Nav}" GroupName="nav" Content="🏠   General" IsChecked="True"/>
          <RadioButton x:Name="navProyecto" Style="{StaticResource Nav}" GroupName="nav" Content="🎬   Proyecto"/>
          <RadioButton x:Name="navAudio"    Style="{StaticResource Nav}" GroupName="nav" Content="🔊   Audio"/>
          <RadioButton x:Name="navSubs"     Style="{StaticResource Nav}" GroupName="nav" Content="💬   Subtítulos"/>

          <Border Height="1" Background="#2A2A3C" Margin="6,14,6,14"/>

          <!-- RAMA 2: torrent + subida al tracker -->
          <TextBlock x:Name="lblSecTorrent" Text="② TORRENT Y SUBIDA" Foreground="{StaticResource AccentBrush}" FontSize="10.5"
                     FontWeight="Bold" Margin="10,0,0,7"/>
          <RadioButton x:Name="navSubida"   Style="{StaticResource Nav}" GroupName="nav" Content="☁   Torrent y subida"/>

          <Border Height="1" Background="#2A2A3C" Margin="6,14,6,14"/>
          <RadioButton x:Name="navAjustes"  Style="{StaticResource Nav}" GroupName="nav" Content="⚙   Ajustes / claves"/>
        </StackPanel>
        <Grid DockPanel.Dock="Bottom" Margin="2,12,2,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock x:Name="lblVersion" Grid.Column="0" VerticalAlignment="Center" Foreground="#56565E"
                     FontSize="11" Margin="11,0,0,0" Text=""/>
          <Button x:Name="btnColapsar" Grid.Column="1" Style="{StaticResource GhostMini}" Content="☰"
                  FontSize="15" Padding="9,3" ToolTip="Ocultar / mostrar el menú"/>
        </Grid>
        <TextBlock x:Name="lblPie" DockPanel.Dock="Bottom" VerticalAlignment="Bottom" Foreground="#56565E" FontSize="11"
                   TextWrapping="Wrap" Margin="13,0,0,10"
                   Text="Las opciones se adaptan a lo detectado en tus vídeos. Todo se decide aquí: el montaje corre sin preguntar nada en consola."/>
        <Border/>
      </DockPanel>
    </Border>

    <!-- ============ CONTENIDO ============ -->
    <Grid Grid.Column="1" Margin="24,18,24,18">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Border Style="{StaticResource Card}" Padding="14,11">
        <StackPanel>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="📁" FontSize="17" VerticalAlignment="Center" Margin="2,0,0,0"/>
            <TextBox x:Name="txtCarpeta" Style="{StaticResource Input}" Grid.Column="1" Margin="11,0,11,0"
                     VerticalAlignment="Center" ToolTip="Carpeta con los vídeos. También puedes arrastrarla a la ventana."/>
            <Button x:Name="btnExaminar" Style="{StaticResource Ghost}" Grid.Column="2" Content="Examinar…"/>
          </Grid>
          <TextBlock x:Name="lblCarpetaDesc" Style="{StaticResource D}" Margin="2,8,0,0"
                     Text="Carpeta con los vídeos a procesar. Las opciones se adaptan a lo que se detecte."/>
          <!-- Aviso: hay originales renombrados «.procesado» en la carpeta. Botón para revertirlos. -->
          <Border x:Name="panProcesado" Visibility="Collapsed" Background="#1C1407" BorderBrush="{StaticResource WarnBrush}"
                  BorderThickness="1" CornerRadius="8" Padding="13,11" Margin="2,11,0,2">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <StackPanel Grid.Column="0" VerticalAlignment="Center">
                <TextBlock x:Name="lblProcesado" Foreground="{StaticResource WarnBrush}" FontWeight="SemiBold" FontSize="13" Text=""/>
                <TextBlock Style="{StaticResource D}" Margin="0,3,0,0"
                           Text="Se renombraron con «.procesado» al conservarse como originales. Quita ese sufijo para volver a trabajar con ellos."/>
              </StackPanel>
              <Button x:Name="btnQuitarProcesado" Grid.Column="1" Style="{StaticResource PrimarySmall}" VerticalAlignment="Center"
                      Margin="12,0,0,0" Content="Quitar «.procesado»"/>
            </Grid>
          </Border>
        </StackPanel>
      </Border>

      <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,6,0">
        <Grid>

          <!-- ============ GENERAL ============ -->
          <StackPanel x:Name="panGeneral">
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Análisis del lote" VerticalAlignment="Center"/>
                  <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock x:Name="lblResumen" Foreground="{StaticResource SubBrush}" FontSize="12" VerticalAlignment="Center" Margin="0,0,12,0"/>
                    <Button x:Name="btnSelTodos" Style="{StaticResource GhostMini}" Content="Todos"/>
                    <Button x:Name="btnSelNinguno" Style="{StaticResource GhostMini}" Content="Ninguno" Margin="6,0,0,0"/>
                  </StackPanel>
                </Grid>
                <!-- Lista redimensionable: arrastra la esquina inferior derecha para verlos todos. -->
                <Grid x:Name="gridListado" Height="250" Margin="0,9,0,0">
                  <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel x:Name="panListado"/>
                  </ScrollViewer>
                  <Thumb x:Name="gripListado" Style="{StaticResource ResizeGrip}" Panel.ZIndex="10"/>
                </Grid>
              </StackPanel>
            </Border>
            <Border x:Name="cardModoLote" Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Modo de procesamiento"/>
                  <TextBlock x:Name="bdgModoLote" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Con varios títulos: ¿comparten proyecto (una temporada) o son títulos distintos?"/>
                <WrapPanel x:Name="chModoLote" Margin="0,11,0,-8"/>
              </StackPanel>
            </Border>
            <!-- Las capturas se configuran SIEMPRE en el panel «Proyecto» (en todos los modos). -->
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Archivos originales al terminar"/>
                <TextBlock Style="{StaticResource D}" Text="Borrarlos ahorra espacio; conservarlos los renombra añadiendo «.procesado»."/>
                <WrapPanel x:Name="chOriginales" Margin="0,11,0,-8"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Sufijo «-HDZ» en el nombre final"/>
                <WrapPanel x:Name="chSufijo" Margin="0,11,0,-8"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Crear archivos .torrent"/>
                <TextBlock Style="{StaticResource D}" Text="Torrents PRIVADOS (no se distribuyen por la red DHT), listos para subir al tracker: uno por archivo, un PACK del lote completo (crea una carpeta con los MKV finales y su torrent multi-archivo), o ambos."/>
                <WrapPanel x:Name="chTorrent" Margin="0,11,0,0"/>
                <StackPanel x:Name="panTorrentDatos" Visibility="Collapsed">
                  <StackPanel x:Name="panPackNombre" Visibility="Collapsed">
                    <TextBlock Style="{StaticResource Lbl}" Text="Nombre del pack (se rellena solo con el nombre del episodio sin el «Exx»; puedes editarlo)"/>
                    <TextBox x:Name="txtPackNombre" Style="{StaticResource Input}" Width="520" HorizontalAlignment="Left"/>
                  </StackPanel>
                  <TextBlock Style="{StaticResource D}" Margin="0,4,0,0" Text="La URL de anuncio del tracker se configura una sola vez en «Ajustes / claves»."/>
                </StackPanel>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Carpetas de salida (opcional)"/>
                <TextBlock Style="{StaticResource D}" Text="Dónde dejar el MKV final y el .torrent. Si lo dejas vacío, se guardan junto a los vídeos de origen. Las capturas SIEMPRE se quedan en la carpeta de origen (la subida las localiza ahí)."/>
                <TextBlock Style="{StaticResource Lbl}" Text="Archivo procesado (MKV final)"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBox x:Name="txtSalidaArchivo" Style="{StaticResource Input}"/>
                  <Button x:Name="btnSalidaArchivo" Style="{StaticResource Ghost}" Grid.Column="1" Margin="8,0,0,0" Content="Examinar…"/>
                </Grid>
                <TextBlock x:Name="lblSalidaTorrent" Style="{StaticResource Lbl}" Text="Archivo .torrent"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBox x:Name="txtSalidaTorrent" Style="{StaticResource Input}"/>
                  <Button x:Name="btnSalidaTorrent" Style="{StaticResource Ghost}" Grid.Column="1" Margin="8,0,0,0" Content="Examinar…"/>
                </Grid>
              </StackPanel>
            </Border>
            <Border x:Name="cardReprocesar" Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Reprocesar archivos ya procesados (-HDZ)"/>
                  <TextBlock x:Name="bdgReprocesar" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Solo aplica si la carpeta no tiene vídeos nuevos pero sí archivos con el sufijo -HDZ."/>
                <WrapPanel x:Name="chReprocesar" Margin="0,11,0,-8"/>
              </StackPanel>
            </Border>
          </StackPanel>

          <!-- ============ PROYECTO ============ -->
          <StackPanel x:Name="panProyecto" Visibility="Collapsed">
            <!-- Datos de proyecto SIEMPRE activos: la GUI los define y los envía, sin prompts en consola.
                 El switch se conserva (oculto y marcado) por compatibilidad con la lógica existente. -->
            <Border Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <CheckBox x:Name="swProyecto" Style="{StaticResource Switch}" IsChecked="True" Content="Definir aquí los datos del proyecto"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <TextBlock Style="{StaticResource D}" Margin="0"
                         Text="El título y el año se usan para nombrar el resultado. El título es obligatorio: el montaje no pregunta nada en consola."/>
            </Border>
            <!-- Barra de pestañas (una por película) — solo en modo «Cada archivo distinto».
                 Réplica de la tira de pestañas de «Torrent y subida»: cada pestaña tiene su propia
                 identidad (título, año, origen, plataforma) y el nombre completo del archivo. -->
            <StackPanel x:Name="panProyTabs" Visibility="Collapsed">
              <TextBlock Style="{StaticResource D}" Margin="2,0,0,8" Text="Cada pestaña es una película distinta del lote: dale a cada una su propio título, año, origen y plataforma. El nombre de la pestaña es el archivo."/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="btnTabProyIzq" Style="{StaticResource GhostMini}" Grid.Column="0" Content="◀" Margin="0,0,3,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
                <ScrollViewer x:Name="scrTabsProy" Grid.Column="1" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Disabled" VerticalAlignment="Bottom">
                  <StackPanel x:Name="panTabsProy" Orientation="Horizontal" HorizontalAlignment="Left"/>
                </ScrollViewer>
                <Button x:Name="btnTabProyDer" Style="{StaticResource GhostMini}" Grid.Column="2" Content="▶" Margin="3,0,0,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
              </Grid>
            </StackPanel>
            <StackPanel x:Name="panProyectoDatos">
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource H}" Text="Identificación"/>
                  <TextBlock x:Name="lblSugerencia" Style="{StaticResource D}" Foreground="{StaticResource WarnBrush}" Visibility="Collapsed"/>
                  <TextBlock Style="{StaticResource Lbl}" Text="Título limpio"/>
                  <TextBox x:Name="txtTitulo" Style="{StaticResource Input}" MaxWidth="520" HorizontalAlignment="Left" Width="520"/>
                  <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                    <StackPanel>
                      <TextBlock Style="{StaticResource Lbl}" Text="Año"/>
                      <TextBox x:Name="txtAno" Style="{StaticResource Input}" Width="120"/>
                    </StackPanel>
                    <CheckBox x:Name="swSerie" Style="{StaticResource Switch}" Content="Es una serie (detecta SxxEyy del nombre)"
                              Margin="30,0,0,6" VerticalAlignment="Bottom"/>
                  </StackPanel>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource H}" Text="Origen"/>
                  <WrapPanel x:Name="chOrigen" Margin="0,11,0,0"/>
                  <StackPanel x:Name="panWeb">
                    <TextBlock Style="{StaticResource Lbl}" Text="Tipo de WEB"/>
                    <WrapPanel x:Name="chWebTipo" Margin="0,0,0,-8"/>
                    <Grid Margin="0,2,0,0">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="250"/><ColumnDefinition Width="18"/><ColumnDefinition Width="250"/>
                      </Grid.ColumnDefinitions>
                      <StackPanel>
                        <TextBlock Style="{StaticResource Lbl}" Text="Plataforma"/>
                        <ComboBox x:Name="cmbPlataforma" Style="{StaticResource Combo}"/>
                      </StackPanel>
                      <StackPanel Grid.Column="2">
                        <TextBlock Style="{StaticResource Lbl}" Text="…o escribe otra etiqueta (opcional)"/>
                        <TextBox x:Name="txtPlataformaOtra" Style="{StaticResource Input}"/>
                      </StackPanel>
                    </Grid>
                  </StackPanel>
                  <StackPanel x:Name="panFisico" Visibility="Collapsed">
                    <Grid Margin="0,2,0,0">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="250"/><ColumnDefinition Width="18"/><ColumnDefinition Width="250"/>
                      </Grid.ColumnDefinitions>
                      <StackPanel>
                        <TextBlock Style="{StaticResource Lbl}" Text="Formato físico"/>
                        <ComboBox x:Name="cmbFormato" Style="{StaticResource Combo}"/>
                      </StackPanel>
                      <StackPanel Grid.Column="2">
                        <TextBlock Style="{StaticResource Lbl}" Text="…o escribe otro formato (opcional)"/>
                        <TextBox x:Name="txtFormatoOtro" Style="{StaticResource Input}"/>
                      </StackPanel>
                    </Grid>
                    <TextBlock Style="{StaticResource Lbl}" Text="Etiquetas extra (opcional)"/>
                    <TextBox x:Name="txtEtiquetas" Style="{StaticResource Input}" Width="518" HorizontalAlignment="Left"/>
                  </StackPanel>
                </StackPanel>
              </Border>
              <!-- Capturas POR PELÍCULA: solo visible en modo «Cada archivo distinto» (cada pestaña
                   puede tener su propio número de capturas). En modo homogéneo se usa el de «General». -->
              <Border x:Name="cardCapturasProy" Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource H}" Text="Capturas por archivo"/>
                  <TextBlock Style="{StaticResource D}" Text="Imágenes JPG del resultado final (con tonemapping si el vídeo es HDR). En «Cada archivo distinto», cada película puede tener su propio número."/>
                  <WrapPanel x:Name="chCapturasProy" Margin="0,11,0,-8"/>
                </StackPanel>
              </Border>
              <!-- En modo heterogéneo aplicamos estos datos a todos los archivos (sin preguntar por
                   archivo en consola). El switch se conserva oculto y marcado. -->
              <Border Style="{StaticResource Card}" Visibility="Collapsed">
                <StackPanel>
                  <CheckBox x:Name="swAplicarTodos" Style="{StaticResource Switch}" Margin="0,12,0,0" IsChecked="True"
                            Content="Aplicar estos datos a TODOS los archivos sin preguntar"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </StackPanel>

          <!-- ============ AUDIO ============ -->
          <StackPanel x:Name="panAudio" Visibility="Collapsed">
            <!-- Pestañas por película (modo «Cada archivo distinto»): el audio se configura por peli. -->
            <StackPanel x:Name="panProyTabsA" Visibility="Collapsed">
              <TextBlock Style="{StaticResource D}" Margin="2,0,0,8" Text="Configurando el AUDIO de cada película por separado. Cambia de pestaña para otra."/>
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <Button x:Name="btnTabProyAIzq" Style="{StaticResource GhostMini}" Grid.Column="0" Content="◀" Margin="0,0,3,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
                <ScrollViewer x:Name="scrTabsProyA" Grid.Column="1" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Disabled" VerticalAlignment="Bottom">
                  <StackPanel x:Name="panTabsProyA" Orientation="Horizontal" HorizontalAlignment="Left"/>
                </ScrollViewer>
                <Button x:Name="btnTabProyADer" Style="{StaticResource GhostMini}" Grid.Column="2" Content="▶" Margin="3,0,0,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
              </Grid>
            </StackPanel>
            <Border x:Name="cardConvAudio" Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Convertir audio"/>
                  <TextBlock x:Name="bdgConvAudio" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Elige una pista y conviértela al formato que quieras (DD+, DD, AAC). Las opciones se adaptan a los canales de la pista (2.0, 5.1, 7.1). Puedes añadir varias conversiones; con «Mantener original» decides si la pista de origen se conserva o se sustituye. El delay (sincronía) se preserva."/>
                <StackPanel x:Name="panConvFilas" Margin="0,12,0,0"/>
                <Button x:Name="btnConvAdd" Style="{StaticResource Ghost}" Content="+ Añadir conversión" HorizontalAlignment="Left" Margin="0,4,0,0"/>
              </StackPanel>
            </Border>
            <Border x:Name="cardDefAudio" Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Pista de audio predeterminada"/>
                  <TextBlock x:Name="bdgDefAudio" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Cuando el idioma principal tiene audio Dolby Y DTS, ¿cuál se marca como pista por defecto? (El orden de pistas no cambia.)"/>
                <WrapPanel x:Name="chDefAudio" Margin="0,11,0,-8"/>
              </StackPanel>
            </Border>
            <Border x:Name="cardUndAudio" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Audios sin idioma definido"/>
                  <TextBlock x:Name="bdgUndAudio" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Estas pistas de audio no declaran idioma. Asigna el de cada una, o déjalas como indeterminado (und)."/>
                <StackPanel x:Name="panUndAudio" Margin="0,10,0,0"/>
              </StackPanel>
            </Border>
            <Border x:Name="phAudio" Style="{StaticResource Card}" Visibility="Collapsed">
              <TextBlock Style="{StaticResource D}" Margin="0"
                         Text="✓  Nada que configurar aquí para este lote: sin audio DTS, sin convivencia Dolby+DTS y todas las pistas con idioma definido."/>
            </Border>
          </StackPanel>

          <!-- ============ SUBTÍTULOS ============ -->
          <StackPanel x:Name="panSubs" Visibility="Collapsed">
            <!-- Pestañas por película (modo «Cada archivo distinto»): los subtítulos se configuran por peli. -->
            <StackPanel x:Name="panProyTabsS" Visibility="Collapsed">
              <TextBlock Style="{StaticResource D}" Margin="2,0,0,8" Text="Configurando los SUBTÍTULOS de cada película por separado. Cambia de pestaña para otra."/>
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <Button x:Name="btnTabProySIzq" Style="{StaticResource GhostMini}" Grid.Column="0" Content="◀" Margin="0,0,3,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
                <ScrollViewer x:Name="scrTabsProyS" Grid.Column="1" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Disabled" VerticalAlignment="Bottom">
                  <StackPanel x:Name="panTabsProyS" Orientation="Horizontal" HorizontalAlignment="Left"/>
                </ScrollViewer>
                <Button x:Name="btnTabProySDer" Style="{StaticResource GhostMini}" Grid.Column="2" Content="▶" Margin="3,0,0,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
              </Grid>
            </StackPanel>
            <Border x:Name="cardUndSub" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Subtítulos sin idioma definido"/>
                  <TextBlock x:Name="bdgUndSub" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Estos subtítulos no declaran idioma. Asigna el de cada uno, o déjalos como indeterminado (und)."/>
                <StackPanel x:Name="panUndSub" Margin="0,10,0,0"/>
              </StackPanel>
            </Border>
            <Border x:Name="cardFiltro" Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Idiomas de subtítulos a mantener"/>
                  <TextBlock x:Name="bdgFiltro" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Quitar un idioma elimina todos sus subtítulos (forzados y completos). Los audios no se tocan."/>
                <WrapPanel x:Name="chFiltro" Margin="0,11,0,0"/>
                <WrapPanel x:Name="wrapIdiomas" Margin="0,6,0,-8" Visibility="Collapsed"/>
              </StackPanel>
            </Border>
            <Border x:Name="cardSubsUnicos" Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Subtítulos sin señal: ¿forzado o completo?"/>
                  <TextBlock x:Name="bdgSubsUnicos" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Estos subtítulos aparecen solos en su idioma y nada indica si son forzados o completos. Decide cada uno:"/>
                <StackPanel x:Name="panSubsUnicos" Margin="0,11,0,0"/>
              </StackPanel>
            </Border>
            <Border x:Name="cardPGS" Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Subtítulos PGS (OCR con Subtitle Edit)"/>
                  <TextBlock x:Name="bdgPGS" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Si se extraen, el script PAUSARÁ en la consola para que los conviertas a SRT con Subtitle Edit y continúe al pulsar ENTER."/>
                <TextBlock Style="{StaticResource Lbl}" Text="¿Extraer los PGS a archivos .sup?"/>
                <WrapPanel x:Name="chExtraerPGS"/>
                <TextBlock Style="{StaticResource Lbl}" Text="Tras convertirlos a SRT…"/>
                <WrapPanel x:Name="chConservarPGS" Margin="0,0,0,-8"/>
              </StackPanel>
            </Border>
            <Border x:Name="phSubs" Style="{StaticResource Card}" Visibility="Collapsed">
              <TextBlock Style="{StaticResource D}" Margin="0"
                         Text="✓  Nada que configurar aquí para este lote: sin PGS, sin ambigüedades, idiomas dentro del límite y todos los subtítulos con idioma definido."/>
            </Border>
          </StackPanel>

          <!-- ============ SUBIR AL TRACKER ============ -->
          <StackPanel x:Name="panSubida" Visibility="Collapsed">
            <!-- TIRA DE PESTAÑAS: todo el panel de subida cuelga del cuerpo de abajo -->
            <TextBlock Style="{StaticResource D}" Margin="2,0,0,8" Text="Cada pestaña es una subida independiente (su propio torrent, capturas, descripción e IDs). Prepara varias —p.ej. 3 películas— y súbelas una a una."/>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Button x:Name="btnTabIzq" Style="{StaticResource GhostMini}" Grid.Column="0" Content="◀" Margin="0,0,3,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
              <ScrollViewer x:Name="scrTabs" Grid.Column="1" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Disabled" VerticalAlignment="Bottom">
                <StackPanel x:Name="panTabsSubida" Orientation="Horizontal" HorizontalAlignment="Left"/>
              </ScrollViewer>
              <Button x:Name="btnTabDer" Style="{StaticResource GhostMini}" Grid.Column="2" Content="▶" Margin="3,0,0,0" VerticalAlignment="Bottom" Visibility="Collapsed"/>
              <Button x:Name="btnNuevaSubida" Style="{StaticResource GhostMini}" Grid.Column="3" Content="➕ Nueva" Margin="10,0,0,0" VerticalAlignment="Bottom"/>
            </Grid>
            <!-- CUERPO de la pestaña activa (la barra de pestañas se conecta a su borde superior) -->
            <Border x:Name="panSubidaBody" Background="#0E0E11" BorderBrush="{StaticResource CardBorderBrush}"
                    BorderThickness="1" CornerRadius="0,10,10,10" Padding="14,16,14,14" Margin="0,0,0,12">
              <StackPanel>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Torrent a subir"/>
                <TextBlock Style="{StaticResource D}" Text="Carga un .torrent ya creado, o crea uno a partir de un archivo (o carpeta/pack) que YA tengas procesado. Se localiza el vídeo, se extrae el MediaInfo y las specs, se cargan las capturas que haya al lado y se rellenan los campos."/>
                <TextBlock Style="{StaticResource Lbl}" Text="Cargar un .torrent existente"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBox x:Name="txtTorrentSubir" Style="{StaticResource Input}" VerticalAlignment="Center" IsReadOnly="True"/>
                  <Button x:Name="btnTorrentExaminar" Style="{StaticResource Ghost}" Grid.Column="1" Content="Examinar…" Margin="9,0,0,0"/>
                  <Button x:Name="btnTorrentUltimo" Style="{StaticResource Ghost}" Grid.Column="2" Content="Último creado" Margin="9,0,0,0"/>
                </Grid>
                <TextBlock Style="{StaticResource Lbl}" Text="…o crear el torrent de algo ya procesado"/>
                <StackPanel Orientation="Horizontal">
                  <Button x:Name="btnCrearTorrentArchivo" Style="{StaticResource Ghost}" Content="🧲  De un archivo…"/>
                  <Button x:Name="btnCrearTorrentCarpeta" Style="{StaticResource Ghost}" Margin="9,0,0,0" Content="🧲  De una carpeta (pack)…"/>
                </StackPanel>
                <Border x:Name="panTorrentProgreso" Visibility="Collapsed" Margin="0,11,0,0"
                        Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                        BorderThickness="1" CornerRadius="8" Padding="12,9">
                  <StackPanel>
                    <Grid Margin="0,0,0,6">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                      <TextBlock x:Name="lblTorrentProgreso" Foreground="{StaticResource TextBrush}" FontSize="12" FontWeight="SemiBold" Text="Creando torrent…"/>
                      <TextBlock x:Name="lblTorrentProgresoPct" Grid.Column="1" Foreground="{StaticResource SubBrush}" FontSize="12" FontWeight="SemiBold" Text="0%"/>
                    </Grid>
                    <ProgressBar x:Name="barTorrentProgreso" Height="6" Minimum="0" Maximum="100" Value="0"
                                 Background="#26262C" BorderThickness="0" Foreground="{StaticResource AccentBrush}"/>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="lblTorrentInfo" Style="{StaticResource D}" Foreground="{StaticResource WarnBrush}" Margin="0,9,0,0" Visibility="Collapsed"/>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Datos del torrent"/>
                <TextBlock Style="{StaticResource Lbl}" Text="Título (nombre de la release)"/>
                <TextBox x:Name="upTitulo" Style="{StaticResource Input}"/>
                <Grid Margin="0,2,0,0">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="16"/><ColumnDefinition Width="*"/><ColumnDefinition Width="16"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <StackPanel><TextBlock Style="{StaticResource Lbl}" Text="Categoría"/><ComboBox x:Name="upCategoria" Style="{StaticResource Combo}"/></StackPanel>
                  <StackPanel Grid.Column="2"><TextBlock Style="{StaticResource Lbl}" Text="Tipo"/><ComboBox x:Name="upTipo" Style="{StaticResource Combo}"/></StackPanel>
                  <StackPanel Grid.Column="4"><TextBlock Style="{StaticResource Lbl}" Text="Resolución"/><ComboBox x:Name="upResolucion" Style="{StaticResource Combo}"/></StackPanel>
                </Grid>
                <StackPanel x:Name="panTV" Visibility="Collapsed" Orientation="Horizontal" Margin="0,2,0,0">
                  <StackPanel><TextBlock Style="{StaticResource Lbl}" Text="Nº temporada"/><TextBox x:Name="upTemporada" Style="{StaticResource Input}" Width="120"/></StackPanel>
                  <StackPanel Margin="16,0,0,0"><TextBlock Style="{StaticResource Lbl}" Text="Nº episodio (0 = pack temporada)"/><TextBox x:Name="upEpisodio" Style="{StaticResource Input}" Width="220"/></StackPanel>
                </StackPanel>
                <TextBlock Style="{StaticResource Lbl}" Text="Palabras clave (separadas por comas)"/>
                <TextBox x:Name="upKeywords" Style="{StaticResource Input}"/>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Identificación (TMDB / IMDB / TVDB / MAL)"/>
                  <Button x:Name="btnBuscarTmdb" Style="{StaticResource GhostMini}" Grid.Column="1" Content="🔍  Buscar en TMDB"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Busca por título y año; al elegir un resultado se rellenan los IDs automáticamente. También puedes escribirlos a mano."/>
                <StackPanel x:Name="panResultadosTmdb" Margin="0,8,0,0"/>
                <Grid Margin="0,4,0,0">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <StackPanel><TextBlock Style="{StaticResource Lbl}" Text="TMDB ID"/><TextBox x:Name="upTmdb" Style="{StaticResource Input}"/></StackPanel>
                  <StackPanel Grid.Column="2"><TextBlock Style="{StaticResource Lbl}" Text="IMDB (nº o tt…)"/><TextBox x:Name="upImdb" Style="{StaticResource Input}"/></StackPanel>
                  <StackPanel Grid.Column="4"><TextBlock Style="{StaticResource Lbl}" Text="TVDB ID"/><TextBox x:Name="upTvdb" Style="{StaticResource Input}"/></StackPanel>
                  <StackPanel Grid.Column="6"><TextBlock Style="{StaticResource Lbl}" Text="MAL ID"/><TextBox x:Name="upMal" Style="{StaticResource Input}"/></StackPanel>
                </Grid>
              </StackPanel>
            </Border>

            <Border x:Name="cardLogos" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Logo de TMDB"/>
                  <TextBlock x:Name="bdgLogos" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Logos disponibles en TMDB para este título. Clic en uno para ponerlo como primera línea de la descripción con [img=700]."/>
                <WrapPanel x:Name="panLogos" Margin="0,11,0,0"/>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Capturas para la descripción"/>
                  <TextBlock x:Name="bdgCapturas" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Si no hay capturas junto al vídeo, créalas aquí (📷). Marca las que quieras y pulsa «Subir y añadir»: se suben al host y se insertan en la descripción como [img=350]. Clic en una miniatura para verla a tamaño completo."/>
                <WrapPanel x:Name="panCapturas" Margin="0,11,0,0"/>
                <Grid Margin="0,10,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <Button x:Name="btnCapTodas" Style="{StaticResource GhostMini}" Grid.Column="0" Content="Todas" VerticalAlignment="Center"/>
                  <Button x:Name="btnCapNinguna" Style="{StaticResource GhostMini}" Grid.Column="1" Content="Ninguna" Margin="6,0,0,0" VerticalAlignment="Center"/>
                  <Button x:Name="btnCapGenerar" Style="{StaticResource GhostMini}" Grid.Column="2" Content="📷  Crear capturas" Margin="14,0,0,0" VerticalAlignment="Center"/>
                  <ComboBox x:Name="cmbNumCapturas" Style="{StaticResource Combo}" Grid.Column="3" Width="62" Margin="6,0,0,0" VerticalAlignment="Center"/>
                  <Button x:Name="btnCapAnadir" Style="{StaticResource GhostMini}" Grid.Column="4" Content="Añadir imágenes…" Margin="6,0,0,0" VerticalAlignment="Center"/>
                  <Button x:Name="btnCapInsertar" Style="{StaticResource PrimarySmall}" Grid.Column="6" Content="⬆  Subir y añadir a descripción" VerticalAlignment="Center"/>
                </Grid>
                <Border x:Name="panSubProgreso" Visibility="Collapsed" Margin="0,10,0,0"
                        Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                        BorderThickness="1" CornerRadius="8" Padding="12,9">
                  <StackPanel>
                    <Grid Margin="0,0,0,6">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                      <TextBlock x:Name="lblSubProgreso" Foreground="{StaticResource TextBrush}" FontSize="12" FontWeight="SemiBold" Text="Subiendo imágenes…"/>
                      <TextBlock x:Name="lblSubProgresoPct" Grid.Column="1" Foreground="{StaticResource SubBrush}" FontSize="12" FontWeight="SemiBold" Text="0/0"/>
                    </Grid>
                    <ProgressBar x:Name="barSubProgreso" Height="6" Minimum="0" Maximum="100" Value="0"
                                 Background="#26262C" BorderThickness="0" Foreground="{StaticResource AccentBrush}"/>
                  </StackPanel>
                </Border>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Descripción (BBCode)"/>
                <TextBlock Style="{StaticResource D}" Text="Va centrada (entre [center]…[/center]). El logo, las capturas y la firma se insertan aquí; da formato con los botones o escribe BBCode a mano."/>
                <!-- Pestañas Escribir / Vista previa (alternan en el mismo sitio, como el tracker) -->
                <StackPanel Orientation="Horizontal" Margin="0,8,0,8">
                  <RadioButton x:Name="tabEscribir" Style="{StaticResource Chip}" GroupName="descTab" Content="✎  Escribir" IsChecked="True"/>
                  <RadioButton x:Name="tabPrevia" Style="{StaticResource Chip}" GroupName="descTab" Content="👁  Vista previa"/>
                </StackPanel>
                <StackPanel x:Name="panEdicionTools">
                  <WrapPanel x:Name="panBbToolbar" Margin="0,2,0,4"/>
                  <Button x:Name="btnDescSubir" Style="{StaticResource GhostMini}" HorizontalAlignment="Left" Margin="0,0,0,6" Content="⬆  Subir imagen del PC…"/>
                </StackPanel>
                <Grid x:Name="gridDesc" Height="300">
                  <TextBox x:Name="upDescripcion" Style="{StaticResource Input}"
                           TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" VerticalContentAlignment="Top"/>
                  <Border x:Name="panPreview" Visibility="Collapsed" Background="#0C0C0F"
                          BorderBrush="{StaticResource ChipBorderBrush}" BorderThickness="1" CornerRadius="8" ClipToBounds="True">
                    <FlowDocumentScrollViewer x:Name="fdPreview" Background="Transparent" Foreground="#E8E8EE"
                                              VerticalScrollBarVisibility="Auto" Padding="0" BorderThickness="0"/>
                  </Border>
                  <Thumb x:Name="gripDesc" Style="{StaticResource ResizeGrip}" Panel.ZIndex="10"/>
                </Grid>
              </StackPanel>
            </Border>

            <Border x:Name="cardFirmas" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="Firma"/>
                  <TextBlock x:Name="bdgFirmas" Style="{StaticResource Bdg}" Grid.Column="1"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Imágenes de la carpeta «firmas». La elegida se sube y se añade al final de la descripción (antes del último [/center]) al subir. Clic para elegir; clic de nuevo para quitar."/>
                <WrapPanel x:Name="panFirmas" Margin="0,11,0,0"/>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBlock Style="{StaticResource H}" Text="MediaInfo"/>
                  <Button x:Name="btnMediaRefrescar" Style="{StaticResource GhostMini}" Grid.Column="1" Content="Regenerar"/>
                </Grid>
                <TextBlock Style="{StaticResource D}" Text="Texto completo del archivo (del primer episodio si es una serie/pack). Se envía en el campo MediaInfo."/>
                <Grid x:Name="gridMediaInfo" Height="120" Margin="0,9,0,0">
                  <TextBox x:Name="upMediainfo" Style="{StaticResource Input}" FontFamily="Consolas"
                           TextWrapping="NoWrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" VerticalContentAlignment="Top"/>
                  <Thumb x:Name="gripMedia" Style="{StaticResource ResizeGrip}" Panel.ZIndex="10"/>
                </Grid>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Opciones"/>
                <WrapPanel Margin="0,11,0,0">
                  <CheckBox x:Name="upAnon" Style="{StaticResource Switch}" Content="Anónimo" Margin="0,0,28,10"/>
                  <CheckBox x:Name="upPersonal" Style="{StaticResource Switch}" Content="Personal release" Margin="0,0,28,10"/>
                  <CheckBox x:Name="upAudioEd" Style="{StaticResource Switch}" Content="Audio editado" Margin="0,0,28,10"/>
                  <CheckBox x:Name="upProper" Style="{StaticResource Switch}" Content="Proper" Margin="0,0,28,10"/>
                </WrapPanel>
                <WrapPanel>
                  <CheckBox x:Name="upDV" Style="{StaticResource Switch}" Content="Dolby Vision" Margin="0,0,28,10"/>
                  <CheckBox x:Name="upHDR10P" Style="{StaticResource Switch}" Content="HDR10+" Margin="0,0,28,10"/>
                  <CheckBox x:Name="upHDR10" Style="{StaticResource Switch}" Content="HDR10" Margin="0,0,28,10"/>
                  <CheckBox x:Name="upModQueue" Style="{StaticResource Switch}" Content="Cola de moderación" Margin="0,0,28,10"/>
                </WrapPanel>
                <TextBlock Style="{StaticResource Lbl}" Text="NFO (opcional)"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <TextBox x:Name="upNfo" Style="{StaticResource Input}" VerticalAlignment="Center" IsReadOnly="True"/>
                  <Button x:Name="btnNfoExaminar" Style="{StaticResource Ghost}" Grid.Column="1" Content="Examinar…" Margin="9,0,0,0"/>
                </Grid>
              </StackPanel>
            </Border>

              </StackPanel>
            </Border>
          </StackPanel>

          <!-- ============ AJUSTES / CLAVES ============ -->
          <StackPanel x:Name="panAjustes" Visibility="Collapsed">
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Credenciales del tracker"/>
                <TextBlock Style="{StaticResource D}" Text="Se guardan solo en tu equipo (%APPDATA%). El token del tracker está en HDZERO → perfil → Configuración → Seguridad."/>
                <TextBlock Style="{StaticResource Lbl}" Text="URL base del tracker"/>
                <TextBox x:Name="cfgTrackerUrl" Style="{StaticResource Input}" Width="420" HorizontalAlignment="Left"/>
                <TextBlock Style="{StaticResource Lbl}" Text="Token de la API del tracker (api_token)"/>
                <TextBox x:Name="cfgTrackerToken" Style="{StaticResource Input}" Width="520" HorizontalAlignment="Left"/>
                <TextBlock Style="{StaticResource Lbl}" Text="URL de anuncio del tracker (announce, con tu passkey)"/>
                <TextBox x:Name="txtAnnounce" Style="{StaticResource Input}" Width="520" HorizontalAlignment="Left"/>
                <TextBlock Style="{StaticResource D}" Margin="0,3,0,0" Text="Se usa al crear los .torrent (en el montaje y en «Torrent y subida»). Se guarda y no se vuelve a pedir."/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="Host de imágenes"/>
                <TextBlock Style="{StaticResource D}" Text="freeimage.host no necesita clave. imgbb requiere una clave gratuita (imgbb.com → About → API)."/>
                <WrapPanel x:Name="chHostImg" Margin="0,11,0,0"/>
                <StackPanel x:Name="panImgbbKey" Visibility="Collapsed">
                  <TextBlock Style="{StaticResource Lbl}" Text="Clave API de imgbb"/>
                  <TextBox x:Name="cfgImgbbKey" Style="{StaticResource Input}" Width="420" HorizontalAlignment="Left"/>
                </StackPanel>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Style="{StaticResource H}" Text="TMDB (búsqueda automática)"/>
                <TextBlock Style="{StaticResource D}" Text="De themoviedb.org → Ajustes → API. Vale cualquiera de las dos credenciales: la «Clave de la API» (corta) o el «Token de acceso de lectura» (largo, empieza por eyJ). Se detecta sola."/>
                <TextBlock Style="{StaticResource Lbl}" Text="Clave API o token de lectura de TMDB"/>
                <TextBox x:Name="cfgTmdbKey" Style="{StaticResource Input}" Width="420" HorizontalAlignment="Left"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <TextBlock Style="{StaticResource D}" Margin="0"
                         Text="⚠  Estas claves dan acceso a tu cuenta. No las compartas. Si alguna se filtró, regenérala en el sitio correspondiente."/>
            </Border>
          </StackPanel>

        </Grid>
      </ScrollViewer>

      <StackPanel Grid.Row="2" Margin="0,14,0,0">
        <!-- Barra de progreso del montaje (oculta hasta lanzar) -->
        <Border x:Name="panProgreso" Visibility="Collapsed" Margin="0,0,0,12"
                Background="{StaticResource ChipBrush}" BorderBrush="{StaticResource ChipBorderBrush}"
                BorderThickness="1" CornerRadius="10" Padding="14,10">
          <StackPanel>
            <Grid Margin="0,0,0,7">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock x:Name="lblProgreso" Foreground="{StaticResource TextBrush}" FontSize="12.5"
                         FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Text="Preparando…"/>
              <TextBlock x:Name="lblProgresoPct" Grid.Column="1" Foreground="{StaticResource SubBrush}"
                         FontSize="12.5" FontWeight="SemiBold" Text="0%"/>
            </Grid>
            <ProgressBar x:Name="barProgreso" Height="8" Minimum="0" Maximum="100" Value="0"
                         Background="#26262C" BorderThickness="0" Foreground="{StaticResource AccentBrush}"/>
            <!-- 2ª barra: detalle de la tarea en curso (ensamblado del MKV actual). Oculta si no aplica. -->
            <StackPanel x:Name="panProgreso2" Visibility="Collapsed" Margin="0,9,0,0">
              <Grid Margin="0,0,0,5">
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <TextBlock x:Name="lblProgreso2" Foreground="{StaticResource SubBrush}" FontSize="11.5"
                           FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Text="Ensamblando…"/>
                <TextBlock x:Name="lblProgreso2Pct" Grid.Column="1" Foreground="{StaticResource SubBrush}"
                           FontSize="11.5" FontWeight="SemiBold" Text="0%"/>
              </Grid>
              <ProgressBar x:Name="barProgreso2" Height="6" Minimum="0" Maximum="100" Value="0"
                           Background="#26262C" BorderThickness="0" Foreground="#6E6E78"/>
            </StackPanel>
          </StackPanel>
        </Border>

        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Button x:Name="btnReset" Style="{StaticResource Ghost}" Grid.Column="0" Margin="0,0,10,0"
                  Content="🗑  Empezar de Zero" ToolTip="Borra la carpeta seleccionada y todas las opciones, dejando el programa como recién abierto."/>
          <TextBlock x:Name="lblEstado" Grid.Column="1" VerticalAlignment="Center" TextWrapping="Wrap"
                     Foreground="{StaticResource SubBrush}" FontSize="12.5" Margin="2,0,16,0"/>
          <Button x:Name="btnConsola" Style="{StaticResource Ghost}" Grid.Column="2" Margin="0,0,10,0"
                  IsEnabled="False"
                  ToolTip="Muestra u oculta la ventana de consola donde corre el montaje (útil si el script hace alguna pregunta).">
            <TextBlock><Run FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" Text="&#xE7B3;"/><Run Text="  Mostrar consola"/></TextBlock>
          </Button>
          <Button x:Name="btnIniciar" Style="{StaticResource Primary}" Grid.Column="3" Content="▶   Iniciar montaje"/>
          <!-- En «Torrent y subida» este sustituye a «Iniciar montaje» (misma celda) -->
          <Button x:Name="btnSubir" Style="{StaticResource Primary}" Grid.Column="3" Content="☁   Subir a HDZERO" Visibility="Collapsed"/>
        </Grid>
      </StackPanel>
    </Grid>
  </Grid>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Parse($xaml)

# AppUserModelID propio: desacopla la ventana del proceso pwsh.exe para que la barra de tareas
# muestre NUESTRO icono (y no el de PowerShell) y agrupe la app por separado. Debe fijarse antes
# de crear el HWND de la ventana (antes de ShowDialog).
try {
    Add-Type -Namespace HdzNative -Name Shell -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("shell32.dll", SetLastError=true)]
public static extern void SetCurrentProcessExplicitAppUserModelID([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string AppID);
'@
    [HdzNative.Shell]::SetCurrentProcessExplicitAppUserModelID("HDZ.Studio")
} catch {}

# =========================================================================
# REFERENCIAS A CONTROLES
# =========================================================================
$refs = @("txtCarpeta","btnExaminar","lblResumen","panListado","imgLogo","panLogoTexto",
          "btnSelTodos","btnSelNinguno","panProcesado","lblProcesado","btnQuitarProcesado",
          "sidebar","btnColapsar","lblVersion","lblSecMontaje","lblSecTorrent","lblPie","panSubidaBody","lblCarpetaDesc",
          "navGeneral","navProyecto","navAudio","navSubs",
          "panGeneral","panProyecto","panAudio","panSubs",
          "chModoLote","chOriginales","chSufijo","chReprocesar",
          "chTorrent","panTorrentDatos","txtAnnounce","panPackNombre","txtPackNombre",
          "txtSalidaArchivo","btnSalidaArchivo","txtSalidaTorrent","btnSalidaTorrent","lblSalidaTorrent",
          "bdgModoLote","bdgReprocesar","bdgConvAudio","bdgDefAudio","bdgUndAudio","bdgUndSub",
          "bdgFiltro","bdgSubsUnicos","bdgPGS","lblSugerencia",
          "cardModoLote","cardReprocesar","cardConvAudio","panConvFilas","btnConvAdd","cardDefAudio","cardUndAudio","cardUndSub",
          "cardFiltro","cardSubsUnicos","cardPGS","phAudio","phSubs","panUndAudio","panUndSub",
          "swProyecto","panProyectoDatos","txtTitulo","txtAno","swSerie",
          "panProyTabs","panTabsProy","scrTabsProy","btnTabProyIzq","btnTabProyDer","cardCapturasProy","chCapturasProy",
          "panProyTabsA","panTabsProyA","scrTabsProyA","btnTabProyAIzq","btnTabProyADer",
          "panProyTabsS","panTabsProyS","scrTabsProyS","btnTabProySIzq","btnTabProySDer",
          "chOrigen","panWeb","chWebTipo","cmbPlataforma","txtPlataformaOtra",
          "panFisico","cmbFormato","txtFormatoOtro","txtEtiquetas","swAplicarTodos",
          "chDefAudio",
          "chFiltro","wrapIdiomas","panSubsUnicos","chExtraerPGS","chConservarPGS",
          "lblEstado","btnIniciar","btnReset","btnConsola",
          "panProgreso","barProgreso","lblProgreso","lblProgresoPct",
          "panProgreso2","barProgreso2","lblProgreso2","lblProgreso2Pct","gridListado","gripListado",
          "navSubida","navAjustes","panSubida","panAjustes",
          "txtTorrentSubir","btnTorrentExaminar","btnTorrentUltimo","lblTorrentInfo",
          "btnCrearTorrentArchivo","btnCrearTorrentCarpeta","panTorrentProgreso","barTorrentProgreso","lblTorrentProgreso","lblTorrentProgresoPct",
          "btnNuevaSubida","panTabsSubida","scrTabs","btnTabIzq","btnTabDer",
          "upTitulo","upCategoria","upTipo","upResolucion","panTV","upTemporada","upEpisodio","upKeywords",
          "btnBuscarTmdb","panResultadosTmdb","upTmdb","upImdb","upTvdb","upMal",
          "cardLogos","bdgLogos","panLogos",
          "bdgCapturas","panCapturas","btnCapTodas","btnCapNinguna","btnCapInsertar","btnCapAnadir","btnCapGenerar","cmbNumCapturas",
          "upDescripcion","upMediainfo","btnMediaRefrescar","gridDesc","gripDesc","gridMediaInfo","gripMedia",
          "panBbToolbar","btnDescSubir","cardFirmas","bdgFirmas","panFirmas",
          "tabEscribir","tabPrevia","panEdicionTools","panPreview","fdPreview",
          "panSubProgreso","barSubProgreso","lblSubProgreso","lblSubProgresoPct",
          "upAnon","upPersonal","upAudioEd","upProper","upDV","upHDR10P","upHDR10","upModQueue",
          "upNfo","btnNfoExaminar","btnSubir",
          "cfgTrackerUrl","cfgTrackerToken","chHostImg","panImgbbKey","cfgImgbbKey","cfgTmdbKey")
$ui = @{}
foreach ($r in $refs) { $ui[$r] = $win.FindName($r) }

$script:bc = New-Object System.Windows.Media.BrushConverter
function Brocha($hex) { return $script:bc.ConvertFromString($hex) }

# =========================================================================
# HELPERS DE CHIPS, COMBOS Y PÍLDORAS
# =========================================================================
$script:gruposChips = @{}

function Add-ChipGroup($panel, $nombre, $opciones, $indiceDefecto = 0, [scriptblock]$alCambiar = $null) {
    $lista = @()
    foreach ($o in $opciones) {
        $rb = New-Object System.Windows.Controls.RadioButton
        $rb.Style = $win.Resources["Chip"]
        $rb.Content = $o.T
        $rb.GroupName = $nombre
        if ($alCambiar) { $rb.Add_Checked($alCambiar) }
        [void]$panel.Children.Add($rb)
        $lista += [PSCustomObject]@{ Radio = $rb; Valor = $o.V }
    }
    $script:gruposChips[$nombre] = $lista
    if ($indiceDefecto -ge 0 -and $indiceDefecto -lt $lista.Count) { $lista[$indiceDefecto].Radio.IsChecked = $true }
}

function Get-ChipValor($nombre) {
    foreach ($e in $script:gruposChips[$nombre]) { if ($e.Radio.IsChecked) { return $e.Valor } }
    return $null
}
function Set-ChipValor($nombre, $valor) {
    foreach ($e in $script:gruposChips[$nombre]) {
        if ("$($e.Valor)" -eq "$valor") { $e.Radio.IsChecked = $true; return }
    }
}

function Init-Combo($cmb, $opciones, $indiceDefecto = 0) {
    foreach ($o in $opciones) { [void]$cmb.Items.Add($o.T) }
    $cmb.Tag = $opciones
    if ($indiceDefecto -ge 0 -and $indiceDefecto -lt $opciones.Count) { $cmb.SelectedIndex = $indiceDefecto }
}
function Get-ComboValor($cmb) {
    if ($cmb.SelectedIndex -lt 0) { return $null }
    return $cmb.Tag[$cmb.SelectedIndex].V
}
function Set-ComboValor($cmb, $valor) {
    for ($i = 0; $i -lt $cmb.Tag.Count; $i++) {
        if ("$($cmb.Tag[$i].V)" -eq "$valor") { $cmb.SelectedIndex = $i; return }
    }
}
function Get-CodigoPlataforma($texto) {
    $t = "$texto".Trim()
    if ($t -match "^(\S+)\s+—") { return $Matches[1] }
    return $t
}

# Píldora pequeña para la lista de análisis ("2160p", "DV", "DTS-HD MA spa"...)
function New-Pill($texto, $estilo = "gris") {
    $b = New-Object System.Windows.Controls.Border
    $b.CornerRadius = [System.Windows.CornerRadius]::new(5)
    $b.Padding = [System.Windows.Thickness]::new(7, 1, 7, 2)
    $b.Margin  = [System.Windows.Thickness]::new(0, 3, 5, 0)
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $texto
    $t.FontSize = 10.5
    switch ($estilo) {
        "accent" { $b.Background = Brocha "#3A1518"; $t.Foreground = Brocha "#FF9B9B" }
        "hdr"    { $b.Background = Brocha "#3A2440"; $t.Foreground = Brocha "#E8A8D8" }
        "warn"   { $b.Background = Brocha "#3E3624"; $t.Foreground = Brocha "#F2C94C" }
        default  { $b.Background = Brocha "#222227"; $t.Foreground = Brocha "#9C9CA8" }
    }
    $b.Child = $t
    return $b
}

# =========================================================================
# LÓGICA DE LA INTERFAZ (antes de crear los chips: sus eventos se disparan ya
# durante la construcción al marcar el chip por defecto)
# =========================================================================
function Set-Estado($texto, $tipo = "info") {
    $ui.lblEstado.Text = $texto
    $ui.lblEstado.Foreground = switch ($tipo) {
        "ok"    { $win.Resources["OkBrush"] }
        "error" { $win.Resources["ErrBrush"] }
        default { $win.Resources["SubBrush"] }
    }
}

function Set-Badge($nombre, $texto, $tipo = "muted") {
    $ui[$nombre].Text = $texto
    $ui[$nombre].Foreground = switch ($tipo) {
        "ok"   { $win.Resources["OkBrush"] }
        "warn" { $win.Resources["WarnBrush"] }
        "sub"  { $win.Resources["SubBrush"] }
        default { $win.Resources["MutedBrush"] }
    }
}

function Actualizar-Proyecto {
    if (-not $ui -or -not $ui.panProyectoDatos) { return }
    $on = [bool]$ui.swProyecto.IsChecked
    $ui.panProyectoDatos.IsEnabled = $on
    $ui.panProyectoDatos.Opacity = if ($on) { 1.0 } else { 0.45 }
    $esWeb = ((Get-ChipValor "origen") -ne "FISICO")
    $ui.panWeb.Visibility    = if ($esWeb) { "Visible" }   else { "Collapsed" }
    $ui.panFisico.Visibility = if ($esWeb) { "Collapsed" } else { "Visible" }
}

function Actualizar-Filtro {
    if (-not $ui -or -not $ui.wrapIdiomas) { return }
    $ui.wrapIdiomas.Visibility = if ((Get-ChipValor "filtro") -eq "PERSONALIZADA") { "Visible" } else { "Collapsed" }
}

function Actualizar-Torrent {
    if (-not $ui -or -not $ui.panTorrentDatos) { return }
    $v = Get-ChipValor "torrent"
    $crearTorrent = ("$v" -in @("INDIVIDUAL", "PACK", "AMBOS"))
    $ui.panTorrentDatos.Visibility = if ($crearTorrent) { "Visible" } else { "Collapsed" }
    $ui.panPackNombre.Visibility   = if ("$v" -in @("PACK", "AMBOS")) { "Visible" } else { "Collapsed" }
    # Si no se va a crear torrent, no tiene sentido elegir dónde guardarlo: deshabilita la ubicación.
    if ($ui.txtSalidaTorrent) { $ui.txtSalidaTorrent.IsEnabled = $crearTorrent }
    if ($ui.btnSalidaTorrent) { $ui.btnSalidaTorrent.IsEnabled = $crearTorrent }
    if ($ui.lblSalidaTorrent) { $ui.lblSalidaTorrent.Opacity   = if ($crearTorrent) { 1.0 } else { 0.4 } }
    Rellenar-PackNombre
}

# Último valor que la identificación automática escribió en Título/Año/Serie. Sirve para que, al
# CAMBIAR de carpeta de vídeos, la sugerencia se actualice sola SIN pisar lo que el usuario haya
# editado a mano: solo se sobrescribe si el campo sigue conteniendo el último valor automático
# (o está vacío). Mismo criterio que usa el nombre del pack.
$script:tituloAutoVal = ""
$script:anoAutoVal    = ""
$script:serieAutoVal  = $null   # $null = aún no sugerido automáticamente
# Carpeta cuya identificación ya se aplicó. Si la carpeta escaneada NO coincide con esta, es que el
# usuario ha cambiado de carpeta → se fuerza la reidentificación (aunque el campo tuviera un valor
# heredado de la sesión anterior). Dentro de la MISMA carpeta se respetan las ediciones manuales.
$script:carpetaIdentificada = ""

# Nombre automático del pack: el del primer episodio seleccionado quitándole el "Exx" (S01E01 → S01),
# es decir el nombre del episodio sin el número de capítulo. Coincide con la convención que usa
# HDZnew para nombrar el .torrent del pack.
$script:packNombreAutoVal = ""
function Get-NombrePackAuto {
    $nombres = @($script:ultimoScan | Where-Object { $script:seleccion[$_.Nombre] -ne $false } |
                 ForEach-Object { $_.Nombre } | Sort-Object)
    if ($nombres.Count -eq 0) { return "" }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($nombres[0])
    return (($base -replace '(?i)\b(S\d{1,3})E\d{1,4}(?:-E?\d{1,4})*\b', '$1') -replace '\s+', ' ').Trim()
}
# Rellena el campo del nombre del pack cuando el torrent está en modo PACK/AMBOS. Respeta una
# edición manual: solo escribe si el campo está vacío o aún contiene el último valor automático.
function Rellenar-PackNombre {
    if (-not $ui -or -not $ui.txtPackNombre) { return }
    if ((Get-ChipValor "torrent") -notin @("PACK", "AMBOS")) { return }
    $auto = Get-NombrePackAuto
    if ([string]::IsNullOrWhiteSpace($auto)) { return }
    $actual = "$($ui.txtPackNombre.Text)"
    if ([string]::IsNullOrWhiteSpace($actual) -or $actual -eq $script:packNombreAutoVal) {
        $ui.txtPackNombre.Text = $auto
        $script:packNombreAutoVal = $auto
    }
}

# Reconstruye los chips de idiomas del filtro personalizado. $detectados = lista de
# @{Cod;Nom;Count} (idiomas reales del lote) o $null para el catálogo completo.
$script:chipsIdiomas = @()
function Reconstruir-ChipsIdiomas($detectados = $null) {
    $marcadosPrev = @($script:chipsIdiomas | Where-Object { $_.IsChecked } | ForEach-Object { "$($_.Tag)" })
    $ui.wrapIdiomas.Children.Clear()
    $script:chipsIdiomas = @()
    $fuente = if ($detectados) { $detectados } else { @($idiomasFiltro | ForEach-Object { @{Cod=$_.Cod; Nom=$_.Nom; Count=$null} }) }
    foreach ($idi in $fuente) {
        $tb = New-Object System.Windows.Controls.Primitives.ToggleButton
        $tb.Style = $win.Resources["ChipToggle"]
        $tb.Content = if ($idi.Count) { "$($idi.Nom)  ·  $($idi.Count)" } else { "$($idi.Nom)" }
        $tb.Tag = $idi.Cod
        if ($marcadosPrev -contains "$($idi.Cod)") { $tb.IsChecked = $true }
        [void]$ui.wrapIdiomas.Children.Add($tb)
        $script:chipsIdiomas += $tb
    }
}

# =========================================================================
# CONSTRUCCIÓN DE GRUPOS DE OPCIONES
# =========================================================================
# Sin opción «Preguntar»: la GUI decide TODO, el montaje corre sin prompts en consola.
# El primer valor de cada grupo es el predeterminado (índice 0, salvo que se indique otro).
# El chip «modoLote» se registra MÁS ABAJO (tras definir Actualizar-ModoLote), porque su callback
# se dispara al marcar la opción por defecto durante la creación y la función debe existir ya.
# Capturas por archivo (SIEMPRE en el panel «Proyecto», en todos los modos).
Add-ChipGroup $ui.chCapturasProy "capturasProy" @(0,1,2,3,4,5,6,8,10,15,20,30 | ForEach-Object { @{T=$(if ($_ -eq 0) { "Ninguna" } else { "$_" }); V=$_} }) 6
Add-ChipGroup $ui.chOriginales "originales" @(
    @{T="Conservarlos (.procesado)"; V=$false},
    @{T="Borrarlos"; V=$true}
)
Add-ChipGroup $ui.chSufijo "sufijo" @(@{T="Sí"; V=$true}, @{T="No"; V=$false})
Add-ChipGroup $ui.chTorrent "torrent" @(
    @{T="No"; V="NO"},
    @{T="Uno por archivo"; V="INDIVIDUAL"},
    @{T="Pack del lote"; V="PACK"},
    @{T="Ambos"; V="AMBOS"}
) 0 { Actualizar-Torrent }
Add-ChipGroup $ui.chReprocesar "reprocesar" @(@{T="No, cancelar"; V=$false}, @{T="Sí, reprocesar"; V=$true})

Add-ChipGroup $ui.chOrigen "origen" @(
    @{T="WEB (streaming)"; V="WEB"},
    @{T="Físico (disco)"; V="FISICO"}
) 0 { Actualizar-Proyecto }
Add-ChipGroup $ui.chWebTipo "webTipo" @(@{T="WEB-DL"; V="WEB-DL"}, @{T="WEBRip"; V="WEBRip"})

Add-ChipGroup $ui.chDefAudio "defAudio" @(
    @{T="Dolby (DD+ / TrueHD / DD)"; V="DOLBY"},
    @{T="DTS (DTS-HD MA / DTS)"; V="DTS"}
)
Add-ChipGroup $ui.chFiltro "filtro" @(
    @{T="Mantener todos"; V="TODOS"},
    @{T="Solo castellano e inglés"; V="CAST_ENG"},
    @{T="Personalizada"; V="PERSONALIZADA"}
) 0 { Actualizar-Filtro }
Add-ChipGroup $ui.chExtraerPGS "extraerPGS" @(
    @{T="No extraer"; V=$false},
    @{T="Sí, extraer y pausar"; V=$true}
)
Add-ChipGroup $ui.chConservarPGS "conservarPGS" @(
    @{T="Mantener PGS + SRT"; V="CONSERVAR_PGS"},
    @{T="Solo SRT (borrar PGS)"; V="ELIMINAR_PGS"}
)

Reconstruir-ChipsIdiomas
Init-Combo $ui.cmbPlataforma (@($plataformasGui | ForEach-Object { @{T=$_; V=(Get-CodigoPlataforma $_)} }))
Init-Combo $ui.cmbFormato    (@($formatosGui    | ForEach-Object { @{T=$_; V=$_} })) 2

# =========================================================================
# MOTOR DE ANÁLISIS DE VÍDEOS
# =========================================================================
# El trabajo pesado (ffprobe por archivo) corre en un runspace aparte para no
# congelar la interfaz. Devuelve filas con la información cruda de cada vídeo.
$scanWorker = {
    param($rutas)
    $resultado = @()
    foreach ($ruta in $rutas) {
        $fila = @{ Ruta = $ruta; Nombre = (Split-Path $ruta -Leaf); Error = $null; Video = $null; Audios = @(); Subs = @() }
        try {
            $raw = & ffprobe -v error -show_streams -of json "$ruta" 2>$null | Out-String
            $json = $raw | ConvertFrom-Json
            if (-not $json -or -not $json.streams) { throw "ffprobe no devolvió datos" }
            foreach ($s in $json.streams) {
                $tipo = "$($s.codec_type)"
                $lang  = ""
                $title = ""
                if ($s.PSObject.Properties.Name -contains "tags" -and $s.tags) {
                    if ($s.tags.PSObject.Properties.Name -contains "language") { $lang = "$($s.tags.language)" }
                    if ($s.tags.PSObject.Properties.Name -contains "title")    { $title = "$($s.tags.title)" }
                }
                $forced = $false
                if ($s.PSObject.Properties.Name -contains "disposition" -and $s.disposition) {
                    $forced = ("$($s.disposition.forced)" -eq "1")
                    if ("$($s.disposition.attached_pic)" -eq "1") { continue }   # carátulas: ignorar
                }
                if ($tipo -eq "video" -and -not $fila.Video) {
                    $esDV = $false
                    if ($s.PSObject.Properties.Name -contains "side_data_list" -and $s.side_data_list) {
                        foreach ($sd in $s.side_data_list) {
                            if ($sd.PSObject.Properties.Name -contains "dv_profile") { $esDV = $true }
                        }
                    }
                    $fila.Video = @{
                        Codec  = "$($s.codec_name)"
                        Width  = [int]"0$($s.width)"
                        Height = [int]"0$($s.height)"
                        EsDV   = $esDV
                        EsHDR  = ("$($s.color_transfer)" -match "smpte2084|arib-std-b67")
                    }
                } elseif ($tipo -eq "audio") {
                    $canales = [int]"0$($s.channels)"
                    $chLayout = ""
                    if ($s.PSObject.Properties.Name -contains "channel_layout") { $chLayout = "$($s.channel_layout)" }
                    $fila.Audios += @{ Index = [int]"0$($s.index)"; Codec = "$($s.codec_name)"; Profile = "$($s.profile)"; Lang = $lang; Title = $title; Forced = $forced; Channels = $canales; ChannelLayout = $chLayout }
                } elseif ($tipo -eq "subtitle") {
                    $fila.Subs += @{ Index = [int]"0$($s.index)"; Codec = "$($s.codec_name)"; Lang = $lang; Title = $title; Forced = $forced
                                     EsPGS = ("$($s.codec_name)" -match "pgs") }
                }
            }
        } catch {
            $fila.Error = "$($_.Exception.Message)"
        }
        $resultado += $fila
    }
    return $resultado
}

function Get-VideosCarpeta($carpeta) {
    return @(Get-ChildItem -LiteralPath $carpeta -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Extension -match "(?i)^\.(mkv|mp4|ts)$" -and $_.Name -notmatch "_FINAL" } |
             Sort-Object Name)
}

# Sugerencia de título / año / serie a partir del nombre del primer archivo
function Sugerir-Proyecto($nombres) {
    $n = @($nombres | Where-Object { $_ -notmatch "-HDZ" } | Select-Object -First 1)
    if (-not $n) { $n = @($nombres | Select-Object -First 1) }
    if (-not $n) { return $null }
    $base = ([System.IO.Path]::GetFileNameWithoutExtension("$n") -replace "[._]", " ")
    $serie = $false
    $corte = $base.Length
    if ($base -match "(?i)\bS(\d{1,3})\s?E(\d{1,4})\b") {
        $serie = $true
        $idx = $base.IndexOf($Matches[0]); if ($idx -ge 0) { $corte = [Math]::Min($corte, $idx) }
    } elseif ($base -match "(?i)\b(\d{1,2})x(\d{2,4})\b") {
        $serie = $true
        $idx = $base.IndexOf($Matches[0]); if ($idx -ge 0) { $corte = [Math]::Min($corte, $idx) }
    }
    $ano = ""
    if ($base -match "\b(19\d{2}|20\d{2})\b") {
        $ano = $Matches[1]
        $idx = $base.IndexOf($Matches[1]); if ($idx -gt 0) { $corte = [Math]::Min($corte, $idx) }
    }
    $titulo = $base.Substring(0, $corte)
    $titulo = (($titulo -replace "[\(\)\[\]]+", " ") -replace "\s+", " ").Trim().Trim('-').Trim()
    if ([string]::IsNullOrWhiteSpace($titulo)) { return $null }
    return @{ Titulo = $titulo; Ano = $ano; Serie = $serie }
}

# Estado neutro de las adaptaciones (sin análisis): tarjetas genéricas visibles,
# sin badges. Las tarjetas de pistas 'und' solo existen con análisis (listan pistas
# concretas), así que arrancan ocultas.
function Reset-Adaptacion {
    foreach ($b in @("bdgModoLote","bdgReprocesar","bdgConvAudio","bdgDefAudio","bdgUndAudio","bdgUndSub","bdgFiltro","bdgSubsUnicos","bdgPGS")) {
        Set-Badge $b ""
    }
    foreach ($c in @("cardModoLote","cardReprocesar","cardConvAudio","cardDefAudio","cardFiltro","cardSubsUnicos","cardPGS")) {
        $ui[$c].Visibility = "Visible"
    }
    $ui.panConvFilas.Children.Clear(); $script:convRows = @(); $script:convTrackOpts = @()
    foreach ($c in @("cardUndAudio","cardUndSub","phAudio","phSubs")) {
        $ui[$c].Visibility = "Collapsed"
    }
    $ui.panUndAudio.Children.Clear()
    $ui.panUndSub.Children.Clear()
    $ui.panSubsUnicos.Children.Clear()
    $script:undRows = @()
    $script:subsUnicosRows = @()
    $ui.lblSugerencia.Visibility = "Collapsed"
    Reconstruir-ChipsIdiomas
}

# Crea una sección por cada subtítulo único ambiguo: etiqueta (idioma · formato) + elección
# Completo/Forzado para ESE sub. Guarda las filas para leerlas al lanzar.
$script:subsUnicosRows = @()
function Construir-FilasSubsUnicos($lista) {
    $ui.panSubsUnicos.Children.Clear()
    $script:subsUnicosRows = @()
    $n = 0
    foreach ($it in @($lista)) {
        $n++
        $g = New-Object System.Windows.Controls.Grid
        $g.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)
        $c1 = New-Object System.Windows.Controls.ColumnDefinition
        $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = [System.Windows.GridLength]::new(250)
        [void]$g.ColumnDefinitions.Add($c1); [void]$g.ColumnDefinitions.Add($c2)
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $it.Label; $t.VerticalAlignment = "Center"; $t.FontSize = 13
        $t.Foreground = $win.Resources["TextBrush"]; $t.TextTrimming = "CharacterEllipsis"
        $t.Margin = [System.Windows.Thickness]::new(0, 0, 12, 0)
        [void]$g.Children.Add($t)
        $sp = New-Object System.Windows.Controls.StackPanel; $sp.Orientation = "Horizontal"
        [System.Windows.Controls.Grid]::SetColumn($sp, 1)
        $grupo = "subuni_$n"
        $rbC = New-Object System.Windows.Controls.RadioButton
        $rbC.Style = $win.Resources["Chip"]; $rbC.Content = "Completo"; $rbC.GroupName = $grupo; $rbC.IsChecked = $true
        $rbF = New-Object System.Windows.Controls.RadioButton
        $rbF.Style = $win.Resources["Chip"]; $rbF.Content = "Forzado"; $rbF.GroupName = $grupo
        [void]$sp.Children.Add($rbC); [void]$sp.Children.Add($rbF)
        [void]$g.Children.Add($sp)
        [void]$ui.panSubsUnicos.Children.Add($g)
        $script:subsUnicosRows += [PSCustomObject]@{ Cod = $it.Cod; Fmt = $it.Fmt; RbForzado = $rbF }
    }
}

# Construye las filas de asignación de idioma POR PISTA para las tarjetas 'und'.
# $pistas = lista de @{Archivo; Id; Tipo("Audio"/"Sub"); Etiqueta}. Si hay varias,
# añade una fila maestra "aplicar a todas".
$script:undRows = @()
function Construir-FilasUnd($panel, $pistas, $tipo) {
    $panel.Children.Clear()
    $nuevaFila = {
        param($etiqueta, $esMaestra)
        $g = New-Object System.Windows.Controls.Grid
        $g.Margin = [System.Windows.Thickness]::new(0, 0, 0, 7)
        $c1 = New-Object System.Windows.Controls.ColumnDefinition
        $c2 = New-Object System.Windows.Controls.ColumnDefinition
        $c2.Width = [System.Windows.GridLength]::new(250)
        [void]$g.ColumnDefinitions.Add($c1); [void]$g.ColumnDefinitions.Add($c2)
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $etiqueta
        $t.FontSize = 12
        $t.VerticalAlignment = "Center"
        $t.TextTrimming = "CharacterEllipsis"
        $t.Margin = [System.Windows.Thickness]::new(0, 0, 12, 0)
        $t.Foreground = if ($esMaestra) { $win.Resources["SubBrush"] } else { $win.Resources["TextBrush"] }
        if ($esMaestra) { $t.FontStyle = [System.Windows.FontStyles]::Italic }
        [void]$g.Children.Add($t)
        $cmb = New-Object System.Windows.Controls.ComboBox
        $cmb.Style = $win.Resources["Combo"]
        [System.Windows.Controls.Grid]::SetColumn($cmb, 1)
        Init-Combo $cmb $opcionesIdiomaUnd
        [void]$g.Children.Add($cmb)
        [void]$panel.Children.Add($g)
        return $cmb
    }
    if ($pistas.Count -gt 1) {
        $cmbMaestra = & $nuevaFila "Aplicar a todas las pistas de abajo:" $true
        $tipoCap = $tipo
        $cmbMaestra.Add_SelectionChanged({
            $idx = $cmbMaestra.SelectedIndex
            foreach ($r in $script:undRows) {
                if ($r.Tipo -eq $tipoCap) { $r.Combo.SelectedIndex = $idx }
            }
        }.GetNewClosure())
    }
    foreach ($p in $pistas) {
        $cmb = & $nuevaFila $p.Etiqueta $false
        $script:undRows += [PSCustomObject]@{ Archivo = $p.Archivo; Id = $p.Id; Tipo = $tipo; Combo = $cmb }
    }
}

# =========================================================================
# CONVERSOR DE AUDIO MANUAL (por pista) — sustituye a la antigua conversión DTS.
# Tarjeta con N filas; cada fila = una pista de origen -> un formato destino, con
# su interruptor «Mantener original». El segundo selector es DINÁMICO según los
# canales de la pista (2.0/5.1/7.1). Estado en $script:convRows; se serializa a
# $cfg.ConversionesAudioManual y se guarda por pestaña en modo heterogéneo.
# =========================================================================
$script:convRows = @()        # filas activas (objetos UI)
$script:convTrackOpts = @()   # opciones del selector de origen (pistas del archivo representativo)

# Nombre comercial del códec de una pista de origen.
function Get-CodecComercial($codec, $profile) {
    switch -Regex ("$codec") {
        "^eac3$"   { return "DD+" }
        "^ac3$"    { return "DD" }
        "^truehd$" { return "TrueHD" }
        "^dts$"    { if ("$profile" -match "MA") { return "DTS-HD MA" } else { return "DTS" } }
        "^aac$"    { return "AAC" }
        "^flac$"   { return "FLAC" }
        "^opus$"   { return "Opus" }
        default    { return ("$codec").ToUpper() }
    }
}
# Nombre de disposición de canales a partir del nº de canales.
function Get-NombreLayout($canales) {
    switch ([int]$canales) {
        1 { return "1.0" }
        2 { return "2.0" }
        6 { return "5.1" }
        8 { return "7.1" }
        default { return "$canales ch" }
    }
}
# Opciones de destino DINÁMICAS según los canales de la pista de origen.
# Devuelve lista de @{T=etiqueta; V="codec|canales|bitrate"}.
function Get-OpcionesDestino($canales) {
    $c = [int]$canales
    # Layouts destino: el mismo que el origen; si es 7.1 (8 ch), también 5.1 (downmix).
    $layouts = @()
    if     ($c -ge 8) { $layouts = @(8, 6) }   # 7.1 -> 7.1 y 5.1
    elseif ($c -ge 6) { $layouts = @(6) }      # 5.1 -> 5.1
    elseif ($c -le 2) { $layouts = @(2) }      # 2.0/1.0 -> 2.0
    else              { $layouts = @($c, 2) }  # raros (5.0, 6.1): su nº y 2.0
    # Bitrate (kbps) por (códec, canales).
    $brFn = {
        param($cod, $ch)
        switch ("$cod") {
            "ac3"  { if ($ch -le 2) { 256 } else { 640 } }
            "eac3" { if ($ch -le 2) { 256 } elseif ($ch -le 6) { 1024 } else { 1280 } }
            "aac"  { if ($ch -le 2) { 256 } elseif ($ch -le 6) { 640 } else { 896 } }
            default { 256 }
        }
    }
    # Familias comerciales; orden DD+, DD, AAC. LÍMITES REALES del codificador ffmpeg:
    # DD+ (eac3) y DD (ac3) NO pasan de 5.1 (6 canales) — el encoder rechaza 7.1. El único
    # destino 7.1 posible con estos formatos es AAC (hasta 8 canales).
    $codecs = @(
        @{ Cod = "eac3"; Nom = "DD+"; MaxCh = 6 },
        @{ Cod = "ac3";  Nom = "DD";  MaxCh = 6 },
        @{ Cod = "aac";  Nom = "AAC"; MaxCh = 8 }
    )
    $ops = @()
    foreach ($ly in $layouts) {
        foreach ($cd in $codecs) {
            if ($ly -gt $cd.MaxCh) { continue }
            $br = & $brFn $cd.Cod $ly
            $ops += @{ T = "$($cd.Nom) $(Get-NombreLayout $ly)"; V = "$($cd.Cod)|$ly|$br" }
        }
    }
    return $ops
}

# Rellena el ComboBox de destino con las opciones válidas para una pista de N canales.
function Llenar-ComboDestino($cmbDest, $canales) {
    $cmbDest.Items.Clear()
    $ops = Get-OpcionesDestino $canales
    foreach ($o in $ops) { [void]$cmbDest.Items.Add($o.T) }
    $cmbDest.Tag = $ops
    if ($ops.Count -gt 0) { $cmbDest.SelectedIndex = 0 }
}

# Añade UNA fila de conversión al panel. $estado opcional = @{Index; Dest; Mant} para restaurar.
function Add-FilaConv($estado = $null) {
    if (@($script:convTrackOpts).Count -eq 0) { return }
    $g = New-Object System.Windows.Controls.Grid
    $g.Margin = [System.Windows.Thickness]::new(0, 0, 0, 9)
    $col0 = New-Object System.Windows.Controls.ColumnDefinition; $col0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = [System.Windows.GridLength]::Auto
    $col2 = New-Object System.Windows.Controls.ColumnDefinition; $col2.Width = [System.Windows.GridLength]::new(185)
    $col3 = New-Object System.Windows.Controls.ColumnDefinition; $col3.Width = [System.Windows.GridLength]::Auto
    $col4 = New-Object System.Windows.Controls.ColumnDefinition; $col4.Width = [System.Windows.GridLength]::Auto
    foreach ($c in @($col0,$col1,$col2,$col3,$col4)) { [void]$g.ColumnDefinitions.Add($c) }

    # Selector de ORIGEN (pista del vídeo).
    $cmbO = New-Object System.Windows.Controls.ComboBox
    $cmbO.Style = $win.Resources["Combo"]
    $cmbO.VerticalAlignment = "Center"
    Init-Combo $cmbO $script:convTrackOpts 0
    [System.Windows.Controls.Grid]::SetColumn($cmbO, 0)
    [void]$g.Children.Add($cmbO)

    # Flecha.
    $flecha = New-Object System.Windows.Controls.TextBlock
    $flecha.Text = "→"; $flecha.FontSize = 15; $flecha.VerticalAlignment = "Center"
    $flecha.Margin = [System.Windows.Thickness]::new(10, 0, 10, 0)
    $flecha.Foreground = $win.Resources["SubBrush"]
    [System.Windows.Controls.Grid]::SetColumn($flecha, 1)
    [void]$g.Children.Add($flecha)

    # Selector de DESTINO (formato), dinámico.
    $cmbD = New-Object System.Windows.Controls.ComboBox
    $cmbD.Style = $win.Resources["Combo"]
    $cmbD.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($cmbD, 2)
    [void]$g.Children.Add($cmbD)

    # Interruptor «Mantener original».
    $sw = New-Object System.Windows.Controls.CheckBox
    $sw.Style = $win.Resources["Switch"]
    $sw.Content = "Mantener original"
    $sw.IsChecked = $true
    $sw.VerticalAlignment = "Center"
    $sw.Margin = [System.Windows.Thickness]::new(16, 0, 6, 0)
    [System.Windows.Controls.Grid]::SetColumn($sw, 3)
    [void]$g.Children.Add($sw)

    # Botón quitar.
    $btnX = New-Object System.Windows.Controls.Button
    $btnX.Style = $win.Resources["GhostMini"]
    $btnX.Content = "✕"
    $btnX.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($btnX, 4)
    [void]$g.Children.Add($btnX)

    # Al cambiar la pista de origen, recalcular las opciones de destino según sus canales.
    $cmbO.Add_SelectionChanged({
        $i = $cmbO.SelectedIndex
        if ($i -lt 0) { return }
        Llenar-ComboDestino $cmbD ([int]$cmbO.Tag[$i].Channels)
    }.GetNewClosure())

    # Quitar la fila.
    $btnX.Add_Click({
        [void]$ui.panConvFilas.Children.Remove($g)
        $script:convRows = @($script:convRows | Where-Object { $_.Grid -ne $g })
    }.GetNewClosure())

    [void]$ui.panConvFilas.Children.Add($g)
    $fila = [PSCustomObject]@{ Grid = $g; CmbOrigen = $cmbO; CmbDestino = $cmbD; SwMant = $sw }
    $script:convRows += $fila

    # Poblar destino una primera vez (la selección de origen ya está en 0).
    Llenar-ComboDestino $cmbD ([int]$script:convTrackOpts[0].Channels)

    # Restaurar estado guardado (pestaña): pista, formato e interruptor.
    if ($estado) {
        for ($k = 0; $k -lt @($script:convTrackOpts).Count; $k++) {
            if ("$($script:convTrackOpts[$k].V)" -eq "$($estado.Index)") { $cmbO.SelectedIndex = $k; break }
        }
        if ($estado.Dest) { Set-ComboValor $cmbD $estado.Dest }
        $sw.IsChecked = [bool]$estado.MantenerOriginal
    }
    return $fila
}

# Reconstruye TODAS las filas de conversión. $estados = lista de @{Index;Dest;Mant} (vacío = 1 fila nueva).
function Reconstruir-FilasConv($estados = @()) {
    $ui.panConvFilas.Children.Clear()
    $script:convRows = @()
    if (@($script:convTrackOpts).Count -eq 0) { return }
    $est = @($estados)
    if ($est.Count -eq 0) { [void](Add-FilaConv) }
    else { foreach ($e in $est) { [void](Add-FilaConv $e) } }
}

# Construye el conversor para un ámbito (homogéneo = selección; heterogéneo = archivos de la peli).
# Toma el PRIMER archivo con audio como representante: en homogéneo todos comparten estructura, y la
# conversión se aplica por índice de pista a cada archivo procesado.
function Construir-Conversor($scope) {
    $rep = @($scope | Where-Object { @($_.Audios).Count -gt 0 } | Select-Object -First 1)
    if (-not $rep) {
        $ui.cardConvAudio.Visibility = "Collapsed"
        $ui.panConvFilas.Children.Clear(); $script:convRows = @(); $script:convTrackOpts = @()
        return
    }
    $script:convTrackOpts = @()
    foreach ($a in $rep.Audios) {
        $lng  = ConvCanon $a.Lang $a.Title
        $nomL = if ($lng -eq "und") { "sin idioma" } else { NombreIdioma $lng }
        $cc   = Get-CodecComercial $a.Codec $a.Profile
        $lay  = Get-NombreLayout $a.Channels
        $script:convTrackOpts += @{ T = "pista $($a.Index) · $nomL · $cc $lay"; V = $a.Index; Channels = $a.Channels; Lang = $lng; Codec = "$($a.Codec)" }
    }
    $ui.cardConvAudio.Visibility = "Visible"
    Reconstruir-FilasConv @()
    Set-Badge "bdgConvAudio" "$(@($rep.Audios).Count) pista(s) en este vídeo" "muted"
}

# Lee las filas de conversión actuales como lista de hashtables (para $cfg / snapshot).
function Leer-ConversionesActuales {
    $lista = @()
    foreach ($r in @($script:convRows)) {
        $idx = Get-ComboValor $r.CmbOrigen
        $dst = Get-ComboValor $r.CmbDestino
        if ($null -eq $idx -or [string]::IsNullOrWhiteSpace("$dst")) { continue }   # fila incompleta: ignorar
        $partes = "$dst" -split "\|"
        if ($partes.Count -lt 3) { continue }
        $opO = @($script:convTrackOpts | Where-Object { "$($_.V)" -eq "$idx" } | Select-Object -First 1)
        $lista += [ordered]@{
            Index            = [int]$idx
            Lang             = if ($opO) { "$($opO.Lang)" } else { "" }
            CodecDestino     = "$($partes[0])"
            CanalesDestino   = [int]$partes[1]
            BitrateK         = [int]$partes[2]
            MantenerOriginal = [bool]$r.SwMant.IsChecked
            Dest             = "$dst"   # "codec|canales|bitrate" para restaurar la fila
        }
    }
    return $lista
}

# Wiring del botón «+ Añadir conversión».
if ($ui.btnConvAdd) { $ui.btnConvAdd.Add_Click({ [void](Add-FilaConv) }) }

# Lista rápida de nombres (feedback inmediato mientras llega el análisis)
function Actualizar-ListaRapida {
    $ui.panListado.Children.Clear()
    $carpeta = $ui.txtCarpeta.Text
    if ([string]::IsNullOrWhiteSpace($carpeta) -or -not (Test-Path -LiteralPath $carpeta)) {
        $ui.lblResumen.Text = "carpeta no válida"
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = "Selecciona o arrastra una carpeta con los vídeos."
        $t.Foreground = $win.Resources["SubBrush"]; $t.FontSize = 12
        [void]$ui.panListado.Children.Add($t)
        return
    }
    $vids = Get-VideosCarpeta $carpeta
    $nuevos = @($vids | Where-Object { $_.Name -notmatch "-HDZ" })
    $hdz    = @($vids | Where-Object { $_.Name -match "-HDZ" })
    $ui.lblResumen.Text = "$($nuevos.Count) por procesar · $($hdz.Count) con -HDZ"
    if ($vids.Count -eq 0) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = "No hay vídeos (.mkv / .mp4 / .ts) en esta carpeta."
        $t.Foreground = $win.Resources["SubBrush"]; $t.FontSize = 12
        [void]$ui.panListado.Children.Add($t)
        return
    }
    foreach ($v in $vids) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = "•  $($v.Name)"
        $t.Foreground = if ($v.Name -match "-HDZ") { $win.Resources["SubBrush"] } else { $win.Resources["TextBrush"] }
        $t.FontSize = 12.5
        $t.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)
        $t.TextTrimming = "CharacterEllipsis"
        [void]$ui.panListado.Children.Add($t)
    }
}

# --- Selección de archivos: qué vídeos del análisis se procesarán ---
$script:seleccion = @{}        # nombre de archivo -> $true/$false (ausente = seleccionado)
$script:ultimoScan = @()       # último resultado de análisis (para re-adaptar sin re-escanear)
$script:carpetaScan = ""       # carpeta a la que pertenece la selección actual
$script:construyendoLista = $false   # evita que los eventos de las casillas se disparen al reconstruir

function Refrescar-Adaptacion {
    if (@($script:ultimoScan).Count -gt 0) { Aplicar-Analisis $script:ultimoScan }
}

# Aplica el resultado del análisis: lista enriquecida (con casillas de selección)
# + adaptación de tarjetas calculada SOLO sobre los archivos seleccionados.
function Aplicar-Analisis($filas) {
    $filas = @($filas | Where-Object { $_ })
    if ($filas.Count -eq 0) { Reset-Adaptacion; return }
    $script:ultimoScan = $filas

    # Entradas de selección nuevas: marcadas por defecto. $sel = subconjunto marcado.
    foreach ($f in $filas) {
        if (-not $script:seleccion.ContainsKey($f.Nombre)) { $script:seleccion[$f.Nombre] = $true }
    }
    $sel = @($filas | Where-Object { $script:seleccion[$_.Nombre] -ne $false })

    # ---- agrupación (mismo criterio que el script) e híbridos ----
    $grupos = [ordered]@{}
    foreach ($f in $filas) {
        $k = (([System.IO.Path]::GetFileNameWithoutExtension($f.Nombre) -replace $normalizadorAgrupacion, "") -replace "\s+", " ").Trim()
        if (-not $grupos.Contains($k)) { $grupos[$k] = @() }
        $grupos[$k] += $f
    }
    $clavesHibridas = @()
    foreach ($k in $grupos.Keys) {
        $g = $grupos[$k]
        if ($g.Count -ge 2) {
            $conDV = @($g | Where-Object { $_.Video -and $_.Video.EsDV })
            $sinDV = @($g | Where-Object { $_.Video -and -not $_.Video.EsDV })
            if ($conDV.Count -gt 0 -and $sinDV.Count -gt 0) { $clavesHibridas += $k }
        }
    }

    # Agrupación SOLO de los seleccionados (la adaptación de tarjetas se basa en esto)
    $gruposSel = [ordered]@{}
    foreach ($f in $sel) {
        $k = (([System.IO.Path]::GetFileNameWithoutExtension($f.Nombre) -replace $normalizadorAgrupacion, "") -replace "\s+", " ").Trim()
        if (-not $gruposSel.Contains($k)) { $gruposSel[$k] = @() }
        $gruposSel[$k] += $f
    }
    # Grupos (= películas) para las pestañas de proyecto del modo heterogéneo: cada grupo aporta
    # su archivo principal (1º) y la lista de archivos que comparten identidad.
    $script:gruposProy = @()
    foreach ($k in $gruposSel.Keys) {
        $archs = @($gruposSel[$k] | ForEach-Object { $_.Nombre })
        if ($archs.Count -gt 0) { $script:gruposProy += @{ Clave = $k; Principal = $archs[0]; Archivos = $archs } }
    }

    # ---- detecciones globales (sobre los SELECCIONADOS) ----
    $archivosDTS  = @($sel | Where-Object { @($_.Audios | Where-Object { $_.Codec -match "dts" }).Count -gt 0 })
    $archivosPGS  = @($sel | Where-Object { @($_.Subs | Where-Object { $_.EsPGS }).Count -gt 0 })
    $pistasUndAudio = @()
    $pistasUndSub   = @()
    $archUndAudio = @{}
    $archUndSub   = @{}
    $idiomasSubs = [ordered]@{}
    foreach ($f in $sel) {
        foreach ($a in $f.Audios) {
            if ((ConvCanon $a.Lang $a.Title) -eq "und") {
                $archUndAudio[$f.Nombre] = $true
                $etCodec = switch -Regex ("$($a.Codec)") {
                    "^eac3$" { "DD+" } "^ac3$" { "DD" } "^truehd$" { "TrueHD" }
                    "^dts$"  { if ("$($a.Profile)" -match "MA") { "DTS-HD MA" } else { "DTS" } }
                    default  { "$($a.Codec)".ToUpper() }
                }
                $etiq = "$($f.Nombre)  —  pista $($a.Index) · $etCodec"
                if ($a.Title) { $etiq += " · «$($a.Title)»" }
                $pistasUndAudio += @{ Archivo = $f.Ruta; Id = $a.Index; Etiqueta = $etiq }
            }
        }
        foreach ($s in $f.Subs) {
            $c = ConvCanon $s.Lang $s.Title
            if ($c -eq "und") {
                $archUndSub[$f.Nombre] = $true
                $fmt = if ($s.EsPGS) { "PGS" } else { "Texto" }
                $etiq = "$($f.Nombre)  —  pista $($s.Index) · $fmt"
                if ($s.Title) { $etiq += " · «$($s.Title)»" }
                $pistasUndSub += @{ Archivo = $f.Ruta; Id = $s.Index; Etiqueta = $etiq }
            } else {
                if (-not $idiomasSubs.Contains($c)) { $idiomasSubs[$c] = 0 }
                $idiomasSubs[$c] = $idiomasSubs[$c] + 1
            }
        }
    }

    # Caso Dolby+DTS en el mismo idioma (por grupo): "SI" directo, "CONV" solo si se
    # convierte DTS (idioma con DTS y sin AC3/EAC3), o "NO".
    $casoDolbyDTS = "NO"
    foreach ($k in $gruposSel.Keys) {
        $porIdioma = @{}
        foreach ($f in $gruposSel[$k]) {
            foreach ($a in $f.Audios) {
                $lng = ConvCanon $a.Lang $a.Title
                if (-not $porIdioma.ContainsKey($lng)) { $porIdioma[$lng] = @{ DTS = $false; Dolby = $false; AC3 = $false } }
                if ($a.Codec -match "dts") { $porIdioma[$lng].DTS = $true }
                if ($a.Codec -match "^(ac3|eac3|truehd)$") { $porIdioma[$lng].Dolby = $true }
                if ($a.Codec -match "^(ac3|eac3)$") { $porIdioma[$lng].AC3 = $true }
            }
        }
        foreach ($lng in $porIdioma.Keys) {
            $i = $porIdioma[$lng]
            if ($i.DTS -and $i.Dolby) { $casoDolbyDTS = "SI" }
            elseif ($i.DTS -and -not $i.AC3 -and $casoDolbyDTS -eq "NO") { $casoDolbyDTS = "CONV" }
        }
        if ($casoDolbyDTS -eq "SI") { break }
    }

    # Subs únicos ambiguos: por grupo, (idioma+formato) con UN solo sub sin señal de forzado/completo.
    # Guardamos un objeto por combinación (clave canónica "cod|Text/PGS") para crear una sección por sub.
    $ambiguosMap = [ordered]@{}
    foreach ($k in $gruposSel.Keys) {
        $combos = @{}
        foreach ($f in $gruposSel[$k]) {
            foreach ($s in $f.Subs) {
                $c = ConvCanon $s.Lang $s.Title
                $fmt = if ($s.EsPGS) { "PGS" } else { "Text" }
                $kk = "$c|$fmt"
                if (-not $combos.ContainsKey($kk)) { $combos[$kk] = @() }
                $combos[$kk] += $s
            }
        }
        foreach ($kk in $combos.Keys) {
            $lista = $combos[$kk]
            if ($lista.Count -ne 1) { continue }
            $s = $lista[0]
            $tieneSenal = $s.Forced -or ("$($s.Title)" -match "(?i)forced|forzado") -or ("$($s.Title)" -match "(?i)completos?|complete|full")
            if (-not $tieneSenal -and -not $ambiguosMap.Contains($kk)) {
                $parts = $kk -split "\|"
                $fmtDisp = if ($parts[1] -eq "PGS") { "PGS" } else { "Texto" }
                $ambiguosMap[$kk] = [PSCustomObject]@{ Cod = $parts[0]; Fmt = $parts[1]; Label = "$(NombreIdioma $parts[0]) · $fmtDisp" }
            }
        }
    }
    $script:subsAmbiguos = @($ambiguosMap.Values)

    $nuevos = @($sel | Where-Object { $_.Nombre -notmatch "-HDZ" })
    $hdz    = @($sel | Where-Object { $_.Nombre -match "-HDZ" })

    # ---- lista enriquecida (con casilla de selección por archivo) ----
    $script:construyendoLista = $true
    $ui.panListado.Children.Clear()
    foreach ($k in $grupos.Keys) {
        foreach ($f in $grupos[$k]) {
            $sp = New-Object System.Windows.Controls.StackPanel
            $sp.Margin = [System.Windows.Thickness]::new(0, 3, 0, 6)
            $marcado = ($script:seleccion[$f.Nombre] -ne $false)
            $sp.Opacity = if ($marcado) { 1.0 } else { 0.45 }
            $tb = New-Object System.Windows.Controls.TextBlock
            $tb.Text = $f.Nombre
            $tb.Foreground = if ($f.Nombre -match "-HDZ") { $win.Resources["SubBrush"] } else { $win.Resources["TextBrush"] }
            $tb.FontSize = 12.5
            $tb.TextTrimming = "CharacterEllipsis"
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Style = $win.Resources["Check"]
            $cb.Content = $tb
            $cb.IsChecked = $marcado
            $cb.Tag = $f.Nombre
            $cb.Add_Click({
                param($s, $e)
                if ($script:construyendoLista) { return }
                $script:seleccion["$($s.Tag)"] = [bool]$s.IsChecked
                Refrescar-Adaptacion
            })
            [void]$sp.Children.Add($cb)
            $wp = New-Object System.Windows.Controls.WrapPanel
            $wp.Margin = [System.Windows.Thickness]::new(26, 0, 0, 0)
            if ($f.Error) {
                [void]$wp.Children.Add((New-Pill "error al analizar: $($f.Error)" "warn"))
            } else {
                if ($f.Video) {
                    $res = if ($f.Video.Height -gt 1200) { "2160p" } elseif ($f.Video.Height -gt 700) { "1080p" } elseif ($f.Video.Height -gt 0) { "$($f.Video.Height)p" } else { "?" }
                    [void]$wp.Children.Add((New-Pill $res "accent"))
                    $cod = switch -Regex ("$($f.Video.Codec)") { "hevc" { "HEVC" } "h264" { "H.264" } "av1" { "AV1" } default { "$($f.Video.Codec)".ToUpper() } }
                    [void]$wp.Children.Add((New-Pill $cod))
                    if ($f.Video.EsDV)  { [void]$wp.Children.Add((New-Pill "Dolby Vision" "hdr")) }
                    if ($f.Video.EsHDR) { [void]$wp.Children.Add((New-Pill "HDR" "hdr")) }
                }
                # resumen de audios: códec comercial + idioma
                $resAud = @()
                foreach ($a in $f.Audios) {
                    $cA = switch -Regex ("$($a.Codec)") {
                        "^eac3$"   { "DD+" } "^ac3$" { "DD" } "^truehd$" { "TrueHD" }
                        "^dts$"    { if ("$($a.Profile)" -match "MA") { "DTS-HD MA" } else { "DTS" } }
                        default    { "$($a.Codec)".ToUpper() }
                    }
                    $lA = ConvCanon $a.Lang $a.Title
                    $resAud += "$cA $(if ($lA -eq 'und') { '¿?' } else { $lA })"
                }
                if ($resAud.Count -gt 0) { [void]$wp.Children.Add((New-Pill ("♪ " + (($resAud | Select-Object -First 4) -join " · ") + $(if ($resAud.Count -gt 4) { " +$($resAud.Count-4)" } else { "" })))) }
                $nPgs = @($f.Subs | Where-Object { $_.EsPGS }).Count
                $nTxt = @($f.Subs | Where-Object { -not $_.EsPGS }).Count
                if ($f.Subs.Count -gt 0) { [void]$wp.Children.Add((New-Pill "💬 $nTxt texto · $nPgs PGS")) }
                if ($clavesHibridas -contains $k) { [void]$wp.Children.Add((New-Pill "PAREJA HÍBRIDA DV+HDR10" "warn")) }
                $undsArchivo = @($f.Audios | Where-Object { (ConvCanon $_.Lang $_.Title) -eq "und" }).Count + @($f.Subs | Where-Object { (ConvCanon $_.Lang $_.Title) -eq "und" }).Count
                if ($undsArchivo -gt 0) { [void]$wp.Children.Add((New-Pill "$undsArchivo pista(s) sin idioma" "warn")) }
            }
            [void]$sp.Children.Add($wp)
            [void]$ui.panListado.Children.Add($sp)
        }
    }
    $script:construyendoLista = $false
    $ui.lblResumen.Text = "$($sel.Count) de $($filas.Count) seleccionado(s) · $($grupos.Keys.Count) título(s) · $($clavesHibridas.Count) híbrido(s)"

    # ---- adaptación de tarjetas: lo que no aplica DESAPARECE (no se atenúa) ----
    # Modo de lote
    if ($gruposSel.Keys.Count -gt 1) {
        $ui.cardModoLote.Visibility = "Visible"
        Set-Badge "bdgModoLote" "$($gruposSel.Keys.Count) títulos distintos seleccionados" "ok"
    } else {
        $ui.cardModoLote.Visibility = "Collapsed"
    }
    # Reprocesar
    if ($nuevos.Count -eq 0 -and $hdz.Count -gt 0) {
        $ui.cardReprocesar.Visibility = "Visible"
        Set-Badge "bdgReprocesar" "solo hay archivos -HDZ en la carpeta" "warn"
    } else {
        $ui.cardReprocesar.Visibility = "Collapsed"
    }
    # Tarjetas de audio/subtítulos: en modo HOMOGÉNEO se construyen aquí del agregado; en
    # HETEROGENEO las construye Aplicar-TabActiva por pestaña (con el archivo de cada película).
    if ((Get-ChipValor "modoLote") -ne "HETEROGENEO") {
    # Conversor de audio (sustituye a la antigua tarjeta DTS)
    Construir-Conversor $sel
    # Default audio
    switch ($casoDolbyDTS) {
        "SI"   { $ui.cardDefAudio.Visibility = "Visible"; Set-Badge "bdgDefAudio" "Dolby y DTS conviven en el mismo idioma" "ok" }
        "CONV" { $ui.cardDefAudio.Visibility = "Visible"; Set-Badge "bdgDefAudio" "aplicará si conviertes DTS a E-AC3" "warn" }
        default { $ui.cardDefAudio.Visibility = "Collapsed" }
    }
    # und audio / und subs: una fila con selector POR PISTA
    if ($pistasUndAudio.Count -gt 0) {
        $ui.cardUndAudio.Visibility = "Visible"
        Construir-FilasUnd $ui.panUndAudio $pistasUndAudio "Audio"
        Set-Badge "bdgUndAudio" "$($pistasUndAudio.Count) pista(s) en $($archUndAudio.Keys.Count) archivo(s)" "warn"
    } else {
        $ui.cardUndAudio.Visibility = "Collapsed"
        $ui.panUndAudio.Children.Clear()
    }
    if ($pistasUndSub.Count -gt 0) {
        $ui.cardUndSub.Visibility = "Visible"
        Construir-FilasUnd $ui.panUndSub $pistasUndSub "Sub"
        Set-Badge "bdgUndSub" "$($pistasUndSub.Count) subtítulo(s) en $($archUndSub.Keys.Count) archivo(s)" "warn"
    } else {
        $ui.cardUndSub.Visibility = "Collapsed"
        $ui.panUndSub.Children.Clear()
    }
    # Filtro de idiomas: chips con los idiomas REALES y sus recuentos
    if ($idiomasSubs.Keys.Count -gt 0) {
        $ui.cardFiltro.Visibility = "Visible"
        $det = @($idiomasSubs.Keys | ForEach-Object { @{Cod=$_; Nom=(NombreIdioma $_); Count=$idiomasSubs[$_]} })
        Reconstruir-ChipsIdiomas $det
        $nota = if ($idiomasSubs.Keys.Count -gt 3) { " — más de 3: revisa el filtro" } else { "" }
        Set-Badge "bdgFiltro" "$($idiomasSubs.Keys.Count) idioma(s): $(@($idiomasSubs.Keys | ForEach-Object { NombreIdioma $_ }) -join ', ')$nota" $(if ($idiomasSubs.Keys.Count -gt 3) { "warn" } else { "ok" })
    } else {
        $ui.cardFiltro.Visibility = "Collapsed"
        Reconstruir-ChipsIdiomas
    }
    # Subs únicos ambiguos: una sección por sub (Construir-FilasSubsUnicos crea una fila por cada uno)
    if (@($script:subsAmbiguos).Count -gt 0) {
        $ui.cardSubsUnicos.Visibility = "Visible"
        Set-Badge "bdgSubsUnicos" "$(@($script:subsAmbiguos).Count) sin definir" "warn"
        Construir-FilasSubsUnicos $script:subsAmbiguos
    } else {
        $ui.cardSubsUnicos.Visibility = "Collapsed"
        $ui.panSubsUnicos.Children.Clear()
        $script:subsUnicosRows = @()
    }
    # PGS
    if ($archivosPGS.Count -gt 0) {
        $ui.cardPGS.Visibility = "Visible"
        $totPgs = 0; foreach ($f in $sel) { $totPgs += @($f.Subs | Where-Object { $_.EsPGS }).Count }
        Set-Badge "bdgPGS" "$totPgs pista(s) PGS en $($archivosPGS.Count) archivo(s)" "ok"
    } else {
        $ui.cardPGS.Visibility = "Collapsed"
    }
    # Tarjetas de "sección sin nada que configurar"
    $audioVacio = ($ui.cardConvAudio.Visibility -ne "Visible") -and ($ui.cardDefAudio.Visibility -ne "Visible") -and ($ui.cardUndAudio.Visibility -ne "Visible")
    $ui.phAudio.Visibility = if ($audioVacio) { "Visible" } else { "Collapsed" }
    $subsVacio = ($ui.cardUndSub.Visibility -ne "Visible") -and ($ui.cardFiltro.Visibility -ne "Visible") -and
                 ($ui.cardSubsUnicos.Visibility -ne "Visible") -and ($ui.cardPGS.Visibility -ne "Visible")
    $ui.phSubs.Visibility = if ($subsVacio) { "Visible" } else { "Collapsed" }
    }  # fin tarjetas modo HOMOGÉNEO

    # Selección vacía: no hay nada que adaptar ni sugerir; avisar y salir.
    if ($sel.Count -eq 0) {
        $ui.phAudio.Visibility = "Collapsed"
        $ui.phSubs.Visibility = "Collapsed"
        $ui.lblSugerencia.Visibility = "Collapsed"
        Set-Estado "Ningún archivo seleccionado: marca al menos uno en la lista para poder procesar." "error"
        return
    }

    # ---- identificación de proyecto ----
    # En modo «Cada archivo distinto» (HETEROGENEO) la identidad es POR pestaña/película:
    # reconstruimos las pestañas (autorrellenando cada una de su nombre) en vez de la identidad única.
    if ((Get-ChipValor "modoLote") -eq "HETEROGENEO") {
        $ui.lblSugerencia.Visibility = "Collapsed"
        Reconstruir-TabsProy
        Set-Estado "Análisis completado: $($sel.Count) de $($filas.Count) vídeo(s), $($gruposSel.Keys.Count) película(s) — revisa cada pestaña."
        Rellenar-PackNombre
        return
    }
    # ---- sugerencia de proyecto (modo homogéneo: una sola identidad para todo el lote) ----
    $sug = Sugerir-Proyecto @($sel | ForEach-Object { $_.Nombre })
    if ($sug) {
        $cambios = @()
        # ¿Es la primera identificación de ESTA carpeta? Si cambiaste de carpeta, se fuerza la
        # reidentificación (sobrescribe aunque el campo tuviera un valor heredado). Dentro de la
        # misma carpeta NO se fuerza: se respeta lo que hayas escrito a mano.
        $forzar = ($script:carpetaScan -ne $script:carpetaIdentificada)
        # Título: sobrescribir si se fuerza, o si está vacío, o si aún contiene la última sugerencia.
        $tActual = "$($ui.txtTitulo.Text)"
        if ($sug.Titulo -and ($forzar -or [string]::IsNullOrWhiteSpace($tActual) -or $tActual -eq $script:tituloAutoVal)) {
            if ($tActual -ne $sug.Titulo) { $cambios += "título" }
            $ui.txtTitulo.Text = $sug.Titulo
            $script:tituloAutoVal = $sug.Titulo
        }
        # Año: mismo criterio. Si la nueva carpeta no aporta año, se limpia el heredado.
        $aActual = "$($ui.txtAno.Text)"
        if ($forzar -or [string]::IsNullOrWhiteSpace($aActual) -or $aActual -eq $script:anoAutoVal) {
            if ($aActual -ne "$($sug.Ano)") { if ($sug.Ano) { $cambios += "año" } }
            $ui.txtAno.Text = "$($sug.Ano)"
            $script:anoAutoVal = "$($sug.Ano)"
        }
        # Serie: se ajusta al forzar, o si no la has tocado desde la última sugerencia.
        if ($forzar -or ($null -eq $script:serieAutoVal) -or ([bool]$ui.swSerie.IsChecked -eq $script:serieAutoVal)) {
            if ([bool]$ui.swSerie.IsChecked -ne [bool]$sug.Serie) { $cambios += "serie" }
            $ui.swSerie.IsChecked = [bool]$sug.Serie
            $script:serieAutoVal  = [bool]$sug.Serie
        }
        # Esta carpeta ya queda identificada: los próximos escaneos de la MISMA carpeta no fuerzan.
        $script:carpetaIdentificada = $script:carpetaScan
        if ($cambios.Count -gt 0) {
            $ui.lblSugerencia.Text = "✦ Identificado del nombre de archivo ($($cambios -join ', ')) — revísalo."
            $ui.lblSugerencia.Visibility = "Visible"
        }
    }

    Set-Estado "Análisis completado: $($sel.Count) de $($filas.Count) vídeo(s) seleccionados, $($gruposSel.Keys.Count) título(s)$(if ($clavesHibridas.Count) { ", $($clavesHibridas.Count) pareja(s) híbrida(s)" })."

    # Autorrellenar el nombre del pack (si está activo el torrent en modo PACK/AMBOS)
    Rellenar-PackNombre
}

# --- Orquestación del escaneo (asíncrono con debounce; síncrono en modo test) ---
$script:scanGen = 0
$script:scanPS = $null
$script:scanAsync = $null
$script:scanGenLanzado = -1

$script:pollTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:pollTimer.Interval = [TimeSpan]::FromMilliseconds(250)
$script:pollTimer.Add_Tick({
    if ($script:scanAsync -and $script:scanAsync.IsCompleted) {
        $script:pollTimer.Stop()
        $res = @()
        try { $res = @($script:scanPS.EndInvoke($script:scanAsync)) } catch {}
        try { $script:scanPS.Dispose() } catch {}
        $script:scanPS = $null; $script:scanAsync = $null
        if ($script:scanGenLanzado -eq $script:scanGen) { Aplicar-Analisis $res }
    }
})

# Detecta originales renombrados «.procesado» en la carpeta actual y muestra/oculta el aviso.
$script:archivosProcesado = @()
function Actualizar-Procesado {
    if (-not $ui -or -not $ui.panProcesado) { return }
    $carpeta = "$($ui.txtCarpeta.Text)".Trim()
    $procesados = @()
    if ($carpeta -and (Test-Path -LiteralPath $carpeta)) {
        try { $procesados = @(Get-ChildItem -LiteralPath $carpeta -File -Filter "*.procesado" -ErrorAction SilentlyContinue) } catch {}
    }
    $script:archivosProcesado = $procesados
    if ($procesados.Count -gt 0) {
        $ui.lblProcesado.Text = "⚠  Hay $($procesados.Count) archivo(s) marcados como «.procesado» en esta carpeta."
        $ui.panProcesado.Visibility = "Visible"
    } else {
        $ui.panProcesado.Visibility = "Collapsed"
    }
}

# =========================================================================
# PESTAÑAS DE PROYECTO (modo «Cada archivo distinto»): una por película, cada una
# con su propia identidad. Mismo patrón que las pestañas de «Torrent y subida»:
# snapshot del estado actual al cambiar de pestaña y restauración del estado destino.
# =========================================================================
$script:tabsProy = @()         # cada elem: @{ Principal=<nombre archivo>; Archivos=@(...); Estado=@{...} }
$script:tabProyActual = -1
$script:gruposProy = @()       # grupos (películas) del último análisis: @{ Clave; Principal; Archivos }

function Nuevo-EstadoProy {
    # Capturas por defecto = el valor global del panel General (chip «capturas»), para no sorprender.
    $capDef = Get-ChipValor "capturasProy"; if ($null -eq $capDef) { $capDef = 6 }
    @{ Titulo=""; Ano=""; EsSerie=$false; Origen="WEB"; WebTipo="WEB-DL"
       PlatIdx=0; PlatOtra=""; FmtIdx=0; FmtOtro=""; Etiquetas=""; Capturas=$capDef }
}
function Snapshot-Proyecto {
    @{
        Titulo  = "$($ui.txtTitulo.Text)"; Ano = "$($ui.txtAno.Text)"; EsSerie = [bool]$ui.swSerie.IsChecked
        Origen  = (Get-ChipValor "origen"); WebTipo = (Get-ChipValor "webTipo")
        PlatIdx = $ui.cmbPlataforma.SelectedIndex; PlatOtra = "$($ui.txtPlataformaOtra.Text)"
        FmtIdx  = $ui.cmbFormato.SelectedIndex;    FmtOtro  = "$($ui.txtFormatoOtro.Text)"
        Etiquetas = "$($ui.txtEtiquetas.Text)"; Capturas = (Get-ChipValor "capturasProy")
    }
}
function Restore-Proyecto($s) {
    $ui.txtTitulo.Text = "$($s.Titulo)"; $ui.txtAno.Text = "$($s.Ano)"; $ui.swSerie.IsChecked = [bool]$s.EsSerie
    if ($s.Origen)  { Set-ChipValor "origen"  $s.Origen }
    if ($s.WebTipo) { Set-ChipValor "webTipo" $s.WebTipo }
    if ([int]$s.PlatIdx -ge 0 -and [int]$s.PlatIdx -lt $ui.cmbPlataforma.Items.Count) { $ui.cmbPlataforma.SelectedIndex = [int]$s.PlatIdx }
    $ui.txtPlataformaOtra.Text = "$($s.PlatOtra)"
    if ([int]$s.FmtIdx -ge 0 -and [int]$s.FmtIdx -lt $ui.cmbFormato.Items.Count) { $ui.cmbFormato.SelectedIndex = [int]$s.FmtIdx }
    $ui.txtFormatoOtro.Text = "$($s.FmtOtro)"
    $ui.txtEtiquetas.Text   = "$($s.Etiquetas)"
    if ($null -ne $s.Capturas) { Set-ChipValor "capturasProy" $s.Capturas }
    Actualizar-Proyecto   # ajusta visibilidad WEB/Físico según el origen restaurado
}
# Las 3 barras de pestañas (Proyecto, Audio, Subtítulos) comparten estado y se dibujan iguales.
function Get-ScrollersProy {
    @($ui.scrTabsProy, $ui.scrTabsProyA, $ui.scrTabsProyS)
}
function Get-PanelesTabsProy {
    @(
        @{ Pan = $ui.panTabsProy;  Scr = $ui.scrTabsProy;  Izq = $ui.btnTabProyIzq;  Der = $ui.btnTabProyDer },
        @{ Pan = $ui.panTabsProyA; Scr = $ui.scrTabsProyA; Izq = $ui.btnTabProyAIzq; Der = $ui.btnTabProyADer },
        @{ Pan = $ui.panTabsProyS; Scr = $ui.scrTabsProyS; Izq = $ui.btnTabProySIzq; Der = $ui.btnTabProySDer }
    )
}
function Update-FlechasTabsProy {
    foreach ($b in Get-PanelesTabsProy) {
        if (-not $b.Scr -or -not $b.Pan) { continue }
        $vp = $b.Scr.ViewportWidth
        if ($vp -le 1) { $b.Izq.Visibility = "Collapsed"; $b.Der.Visibility = "Collapsed"; continue }
        $vis = if ($b.Scr.ExtentWidth -gt ($vp + 1)) { "Visible" } else { "Collapsed" }
        $b.Izq.Visibility = $vis; $b.Der.Visibility = $vis
    }
}
function Construir-TabBarProy {
    foreach ($b in Get-PanelesTabsProy) {
        if (-not $b.Pan) { continue }
        $b.Pan.Children.Clear()
        for ($i = 0; $i -lt @($script:tabsProy).Count; $i++) {
            $idx = $i
            $titulo = "$($script:tabsProy[$i].Principal)"
            if ([string]::IsNullOrWhiteSpace($titulo)) { $titulo = "Película $($i+1)" }
            $cont = New-Object System.Windows.Controls.Border
            $cont.CornerRadius = [System.Windows.CornerRadius]::new(7,7,0,0)
            $cont.Cursor = [System.Windows.Input.Cursors]::Hand
            if ($i -eq $script:tabProyActual) {
                $cont.Background = (Brocha "#0E0E11"); $cont.BorderBrush = $win.Resources["AccentBrush"]
                $cont.BorderThickness = [System.Windows.Thickness]::new(0,2,0,0); $cont.Margin = [System.Windows.Thickness]::new(0,0,3,-1)
                $cont.Padding = [System.Windows.Thickness]::new(13,7,13,8)
            } else {
                $cont.Background = (Brocha "#17171C"); $cont.BorderBrush = $win.Resources["CardBorderBrush"]
                $cont.BorderThickness = [System.Windows.Thickness]::new(1); $cont.Margin = [System.Windows.Thickness]::new(0,4,3,0)
                $cont.Padding = [System.Windows.Thickness]::new(13,5,13,6)
            }
            $t = New-Object System.Windows.Controls.TextBlock
            $t.Text = $titulo; $t.VerticalAlignment = "Center"; $t.FontSize = 12; $t.ToolTip = $titulo
            $t.Foreground = $(if ($i -eq $script:tabProyActual) { [System.Windows.Media.Brushes]::White } else { $win.Resources["SubBrush"] })
            $cont.Child = $t
            $cont.Add_MouseLeftButtonUp({ param($snd, $e) Cambiar-TabProy $idx }.GetNewClosure())
            [void]$b.Pan.Children.Add($cont)
        }
    }
    Update-FlechasTabsProy
}
function Cambiar-TabProy($idx) {
    if ($idx -eq $script:tabProyActual -or $idx -lt 0 -or $idx -ge @($script:tabsProy).Count) { return }
    if ($script:tabProyActual -ge 0 -and $script:tabProyActual -lt @($script:tabsProy).Count) {
        $script:tabsProy[$script:tabProyActual].Estado = Snapshot-Proyecto
        $script:tabsProy[$script:tabProyActual].Decisiones = Snapshot-DecisionesTab
    }
    $script:tabProyActual = $idx
    Restore-Proyecto $script:tabsProy[$idx].Estado
    Aplicar-TabActiva          # reconstruye las tarjetas de audio/subs de esta peli y aplica sus decisiones
    Construir-TabBarProy
}
# Reconstruye las pestañas a partir de los grupos del último análisis ($script:gruposProy),
# conservando lo que el usuario ya hubiera editado (indexado por archivo principal) y
# autorrellenando las pestañas nuevas desde el nombre del archivo.
function Reconstruir-TabsProy {
    if (-not $ui.panTabsProy) { return }
    if ($script:tabProyActual -ge 0 -and $script:tabProyActual -lt @($script:tabsProy).Count) {
        $script:tabsProy[$script:tabProyActual].Estado = Snapshot-Proyecto
        $script:tabsProy[$script:tabProyActual].Decisiones = Snapshot-DecisionesTab
    }
    $previos = @{}
    foreach ($t in @($script:tabsProy)) { if ($t.Principal) { $previos[$t.Principal] = $t } }
    $nuevas = @()
    foreach ($g in @($script:gruposProy)) {
        if ($previos.ContainsKey($g.Principal)) {
            $t = $previos[$g.Principal]; $t.Archivos = $g.Archivos; $nuevas += $t
        } else {
            $est = Nuevo-EstadoProy
            $sug = Sugerir-Proyecto @($g.Principal)
            if ($sug) { $est.Titulo = "$($sug.Titulo)"; $est.Ano = "$($sug.Ano)"; $est.EsSerie = [bool]$sug.Serie }
            # Decisiones = $null → al activar la pestaña se usan los valores por defecto de sus tarjetas.
            $nuevas += @{ Principal = $g.Principal; Archivos = $g.Archivos; Estado = $est; Decisiones = $null }
        }
    }
    $script:tabsProy = @($nuevas)
    if (@($script:tabsProy).Count -eq 0) { $script:tabProyActual = -1; Construir-TabBarProy; return }
    if ($script:tabProyActual -lt 0 -or $script:tabProyActual -ge @($script:tabsProy).Count) { $script:tabProyActual = 0 }
    # Pre-cargar las decisiones por defecto de CADA pestaña (construyendo sus tarjetas una vez), para
    # poder enviar valores correctos aunque el usuario no llegue a abrir esa pestaña (p.ej. pistas und).
    for ($i = 0; $i -lt @($script:tabsProy).Count; $i++) {
        if ($null -eq $script:tabsProy[$i].Decisiones) {
            $nombresI = @($script:tabsProy[$i].Archivos)
            $scopeI = @($script:ultimoScan | Where-Object { $nombresI -contains $_.Nombre })
            Construir-Adaptacion $scopeI
            $script:tabsProy[$i].Decisiones = Snapshot-DecisionesTab
        }
    }
    Restore-Proyecto $script:tabsProy[$script:tabProyActual].Estado
    Aplicar-TabActiva          # deja en pantalla las tarjetas de la pestaña activa
    Construir-TabBarProy
}
# --- FASE 2: tarjetas de audio/subtítulos POR PESTAÑA ---
# Reconstruye SOLO las tarjetas de adaptación (DTS, audio por defecto, pistas und, filtro de
# idiomas, subs únicos, PGS) a partir de un subconjunto de archivos ($selScope = entradas de
# $script:ultimoScan). Es una copia parametrizada de la misma lógica de Aplicar-Analisis, para
# poder mostrar las tarjetas del archivo de la pestaña activa sin tocar el camino homogéneo.
function Construir-Adaptacion($selScope) {
    $selScope = @($selScope | Where-Object { $_ })
    # Construir-FilasUnd ACUMULA en $script:undRows (no lo resetea): al construir varias pestañas
    # seguidas se mezclarían las pistas. Lo limpiamos aquí para que cada ámbito empiece de cero.
    $script:undRows = @()
    $gruposScope = [ordered]@{}
    foreach ($f in $selScope) {
        $k = (([System.IO.Path]::GetFileNameWithoutExtension($f.Nombre) -replace $normalizadorAgrupacion, "") -replace "\s+", " ").Trim()
        if (-not $gruposScope.Contains($k)) { $gruposScope[$k] = @() }
        $gruposScope[$k] += $f
    }
    # Detecciones
    $archivosDTS  = @($selScope | Where-Object { @($_.Audios | Where-Object { $_.Codec -match "dts" }).Count -gt 0 })
    $archivosPGS  = @($selScope | Where-Object { @($_.Subs | Where-Object { $_.EsPGS }).Count -gt 0 })
    $pistasUndAudio = @(); $pistasUndSub = @(); $archUndAudio = @{}; $archUndSub = @{}; $idiomasSubs = [ordered]@{}
    foreach ($f in $selScope) {
        foreach ($a in $f.Audios) {
            if ((ConvCanon $a.Lang $a.Title) -eq "und") {
                $archUndAudio[$f.Nombre] = $true
                $etCodec = switch -Regex ("$($a.Codec)") {
                    "^eac3$" { "DD+" } "^ac3$" { "DD" } "^truehd$" { "TrueHD" }
                    "^dts$"  { if ("$($a.Profile)" -match "MA") { "DTS-HD MA" } else { "DTS" } }
                    default  { "$($a.Codec)".ToUpper() }
                }
                $etiq = "$($f.Nombre)  —  pista $($a.Index) · $etCodec"
                if ($a.Title) { $etiq += " · «$($a.Title)»" }
                $pistasUndAudio += @{ Archivo = $f.Ruta; Id = $a.Index; Etiqueta = $etiq }
            }
        }
        foreach ($s in $f.Subs) {
            $c = ConvCanon $s.Lang $s.Title
            if ($c -eq "und") {
                $archUndSub[$f.Nombre] = $true
                $fmt = if ($s.EsPGS) { "PGS" } else { "Texto" }
                $etiq = "$($f.Nombre)  —  pista $($s.Index) · $fmt"
                if ($s.Title) { $etiq += " · «$($s.Title)»" }
                $pistasUndSub += @{ Archivo = $f.Ruta; Id = $s.Index; Etiqueta = $etiq }
            } else {
                if (-not $idiomasSubs.Contains($c)) { $idiomasSubs[$c] = 0 }
                $idiomasSubs[$c] = $idiomasSubs[$c] + 1
            }
        }
    }
    $casoDolbyDTS = "NO"
    foreach ($k in $gruposScope.Keys) {
        $porIdioma = @{}
        foreach ($f in $gruposScope[$k]) {
            foreach ($a in $f.Audios) {
                $lng = ConvCanon $a.Lang $a.Title
                if (-not $porIdioma.ContainsKey($lng)) { $porIdioma[$lng] = @{ DTS = $false; Dolby = $false; AC3 = $false } }
                if ($a.Codec -match "dts") { $porIdioma[$lng].DTS = $true }
                if ($a.Codec -match "^(ac3|eac3|truehd)$") { $porIdioma[$lng].Dolby = $true }
                if ($a.Codec -match "^(ac3|eac3)$") { $porIdioma[$lng].AC3 = $true }
            }
        }
        foreach ($lng in $porIdioma.Keys) {
            $i = $porIdioma[$lng]
            if ($i.DTS -and $i.Dolby) { $casoDolbyDTS = "SI" }
            elseif ($i.DTS -and -not $i.AC3 -and $casoDolbyDTS -eq "NO") { $casoDolbyDTS = "CONV" }
        }
        if ($casoDolbyDTS -eq "SI") { break }
    }
    $ambiguosMap = [ordered]@{}
    foreach ($k in $gruposScope.Keys) {
        $combos = @{}
        foreach ($f in $gruposScope[$k]) {
            foreach ($s in $f.Subs) {
                $c = ConvCanon $s.Lang $s.Title
                $fmt = if ($s.EsPGS) { "PGS" } else { "Text" }
                $kk = "$c|$fmt"
                if (-not $combos.ContainsKey($kk)) { $combos[$kk] = @() }
                $combos[$kk] += $s
            }
        }
        foreach ($kk in $combos.Keys) {
            $lista = $combos[$kk]
            if ($lista.Count -ne 1) { continue }
            $s = $lista[0]
            $tieneSenal = $s.Forced -or ("$($s.Title)" -match "(?i)forced|forzado") -or ("$($s.Title)" -match "(?i)completos?|complete|full")
            if (-not $tieneSenal -and -not $ambiguosMap.Contains($kk)) {
                $parts = $kk -split "\|"
                $fmtDisp = if ($parts[1] -eq "PGS") { "PGS" } else { "Texto" }
                $ambiguosMap[$kk] = [PSCustomObject]@{ Cod = $parts[0]; Fmt = $parts[1]; Label = "$(NombreIdioma $parts[0]) · $fmtDisp" }
            }
        }
    }
    $script:subsAmbiguos = @($ambiguosMap.Values)
    # Aplicación de tarjetas (idéntico criterio que Aplicar-Analisis, ámbito = $selScope)
    Construir-Conversor $selScope
    switch ($casoDolbyDTS) {
        "SI"   { $ui.cardDefAudio.Visibility = "Visible"; Set-Badge "bdgDefAudio" "Dolby y DTS conviven en el mismo idioma" "ok" }
        "CONV" { $ui.cardDefAudio.Visibility = "Visible"; Set-Badge "bdgDefAudio" "aplicará si conviertes DTS a E-AC3" "warn" }
        default { $ui.cardDefAudio.Visibility = "Collapsed" }
    }
    if ($pistasUndAudio.Count -gt 0) {
        $ui.cardUndAudio.Visibility = "Visible"; Construir-FilasUnd $ui.panUndAudio $pistasUndAudio "Audio"
        Set-Badge "bdgUndAudio" "$($pistasUndAudio.Count) pista(s) en $($archUndAudio.Keys.Count) archivo(s)" "warn"
    } else { $ui.cardUndAudio.Visibility = "Collapsed"; $ui.panUndAudio.Children.Clear() }
    if ($pistasUndSub.Count -gt 0) {
        $ui.cardUndSub.Visibility = "Visible"; Construir-FilasUnd $ui.panUndSub $pistasUndSub "Sub"
        Set-Badge "bdgUndSub" "$($pistasUndSub.Count) subtítulo(s) en $($archUndSub.Keys.Count) archivo(s)" "warn"
    } else { $ui.cardUndSub.Visibility = "Collapsed"; $ui.panUndSub.Children.Clear() }
    if ($idiomasSubs.Keys.Count -gt 0) {
        $ui.cardFiltro.Visibility = "Visible"
        $det = @($idiomasSubs.Keys | ForEach-Object { @{Cod=$_; Nom=(NombreIdioma $_); Count=$idiomasSubs[$_]} })
        Reconstruir-ChipsIdiomas $det
        $nota = if ($idiomasSubs.Keys.Count -gt 3) { " — más de 3: revisa el filtro" } else { "" }
        Set-Badge "bdgFiltro" "$($idiomasSubs.Keys.Count) idioma(s): $(@($idiomasSubs.Keys | ForEach-Object { NombreIdioma $_ }) -join ', ')$nota" $(if ($idiomasSubs.Keys.Count -gt 3) { "warn" } else { "ok" })
    } else { $ui.cardFiltro.Visibility = "Collapsed"; Reconstruir-ChipsIdiomas }
    if (@($script:subsAmbiguos).Count -gt 0) {
        $ui.cardSubsUnicos.Visibility = "Visible"; Set-Badge "bdgSubsUnicos" "$(@($script:subsAmbiguos).Count) sin definir" "warn"
        Construir-FilasSubsUnicos $script:subsAmbiguos
    } else { $ui.cardSubsUnicos.Visibility = "Collapsed"; $ui.panSubsUnicos.Children.Clear(); $script:subsUnicosRows = @() }
    if ($archivosPGS.Count -gt 0) {
        $ui.cardPGS.Visibility = "Visible"
        $totPgs = 0; foreach ($f in $selScope) { $totPgs += @($f.Subs | Where-Object { $_.EsPGS }).Count }
        Set-Badge "bdgPGS" "$totPgs pista(s) PGS en $($archivosPGS.Count) archivo(s)" "ok"
    } else { $ui.cardPGS.Visibility = "Collapsed" }
    $audioVacio = ($ui.cardConvAudio.Visibility -ne "Visible") -and ($ui.cardDefAudio.Visibility -ne "Visible") -and ($ui.cardUndAudio.Visibility -ne "Visible")
    $ui.phAudio.Visibility = if ($audioVacio) { "Visible" } else { "Collapsed" }
    $subsVacio = ($ui.cardUndSub.Visibility -ne "Visible") -and ($ui.cardFiltro.Visibility -ne "Visible") -and
                 ($ui.cardSubsUnicos.Visibility -ne "Visible") -and ($ui.cardPGS.Visibility -ne "Visible")
    $ui.phSubs.Visibility = if ($subsVacio) { "Visible" } else { "Collapsed" }
}
# Snapshot de las decisiones de audio/subs de la pestaña activa (lee las tarjetas actuales).
function Snapshot-DecisionesTab {
    $und = @{}
    foreach ($r in @($script:undRows)) { $und["$($r.Archivo)|$($r.Id)|$($r.Tipo)"] = (Get-ComboValor $r.Combo) }
    $su = @{}
    foreach ($r in @($script:subsUnicosRows)) { $su["$($r.Cod)|$($r.Fmt)"] = $(if ($r.RbForzado.IsChecked) { "Forzado" } else { "Completo" }) }
    @{
        Conv = (Leer-ConversionesActuales); DefAudio = (Get-ChipValor "defAudio"); Filtro = (Get-ChipValor "filtro")
        FiltroLangs = @($script:chipsIdiomas | Where-Object { $_.IsChecked } | ForEach-Object { "$($_.Tag)" })
        ExtraerPGS = (Get-ChipValor "extraerPGS"); ConservarPGS = (Get-ChipValor "conservarPGS")
        Und = $und; SubsUnicos = $su
    }
}
# Aplica las decisiones guardadas sobre las tarjetas YA reconstruidas (Construir-Adaptacion antes).
function Restore-DecisionesTab($d) {
    if (-not $d) { return }
    if ($null -ne $d.Conv) { Reconstruir-FilasConv @($d.Conv) }
    if ($d.DefAudio)     { Set-ChipValor "defAudio" $d.DefAudio }
    if ($d.ExtraerPGS)   { Set-ChipValor "extraerPGS" $d.ExtraerPGS }
    if ($d.ConservarPGS) { Set-ChipValor "conservarPGS" $d.ConservarPGS }
    if ($d.Filtro)       { Set-ChipValor "filtro" $d.Filtro; Actualizar-Filtro }
    if ($d.Filtro -eq "PERSONALIZADA" -and $d.FiltroLangs) {
        foreach ($tb in $script:chipsIdiomas) { $tb.IsChecked = (@($d.FiltroLangs) -contains "$($tb.Tag)") }
    }
    if ($d.Und) {
        foreach ($r in @($script:undRows)) {
            $k = "$($r.Archivo)|$($r.Id)|$($r.Tipo)"
            if ($d.Und.ContainsKey($k) -and $null -ne $d.Und[$k]) { Set-ComboValor $r.Combo $d.Und[$k] }
        }
    }
    if ($d.SubsUnicos) {
        foreach ($r in @($script:subsUnicosRows)) {
            $k = "$($r.Cod)|$($r.Fmt)"
            if ($d.SubsUnicos.ContainsKey($k) -and $d.SubsUnicos[$k] -eq "Forzado") { $r.RbForzado.IsChecked = $true }
        }
    }
}
# Reconstruye las tarjetas de audio/subs para la pestaña activa y aplica sus decisiones guardadas.
function Aplicar-TabActiva {
    if ($script:tabProyActual -lt 0 -or $script:tabProyActual -ge @($script:tabsProy).Count) { return }
    $tab = $script:tabsProy[$script:tabProyActual]
    $nombres = @($tab.Archivos)
    $scope = @($script:ultimoScan | Where-Object { $nombres -contains $_.Nombre })
    Construir-Adaptacion $scope
    Restore-DecisionesTab $tab.Decisiones
}
# Muestra/oculta la barra de pestañas según el modo de lote.
function Actualizar-ModoLote {
    if (-not $ui -or -not $ui.panProyTabs) { return }
    $het = ((Get-ChipValor "modoLote") -eq "HETEROGENEO")
    $visTabs = if ($het) { "Visible" } else { "Collapsed" }
    foreach ($p in @($ui.panProyTabs, $ui.panProyTabsA, $ui.panProyTabsS)) { if ($p) { $p.Visibility = $visTabs } }
    # Capturas: la tarjeta vive SIEMPRE en Proyecto (no se oculta). En heterogéneo el valor es por
    # pestaña; en homogéneo es un único valor que se envía como NumCapturas global.
    if ($het) { Reconstruir-TabsProy }
    Update-FlechasTabsProy
}
# Registro del chip «modoLote» (aquí, ya con Actualizar-ModoLote definida: su callback se dispara
# al fijar la opción por defecto durante la creación).
Add-ChipGroup $ui.chModoLote "modoLote" @(
    @{T="Mismo proyecto (temporada)"; V="HOMOGENEO"},
    @{T="Cada archivo distinto"; V="HETEROGENEO"}
) 0 { Actualizar-ModoLote }

function Lanzar-Escaneo {
    $script:scanGen++
    Reset-Adaptacion
    Actualizar-Procesado
    Actualizar-ListaRapida
    $carpeta = $ui.txtCarpeta.Text
    $script:ultimoScan = @()
    if ([string]::IsNullOrWhiteSpace($carpeta) -or -not (Test-Path -LiteralPath $carpeta)) { return }
    # La selección de archivos pertenece a UNA carpeta: al cambiar de carpeta se reinicia.
    if ($carpeta -ne $script:carpetaScan) {
        $script:seleccion = @{}
        $script:carpetaScan = $carpeta
    }
    $vids = Get-VideosCarpeta $carpeta
    # Analizamos lo que el script procesaría: los nuevos; si no hay, los -HDZ (reproceso).
    $nuevos = @($vids | Where-Object { $_.Name -notmatch "-HDZ" })
    $objetivo = if ($nuevos.Count -gt 0) { $nuevos } else { $vids }
    if ($objetivo.Count -eq 0) { Set-Estado "No hay vídeos que analizar en esta carpeta."; return }

    $rutas = @($objetivo | ForEach-Object { $_.FullName })
    $ui.lblResumen.Text = "⏳ analizando $($rutas.Count) vídeo(s)…"
    Set-Estado "Analizando pistas de $($rutas.Count) vídeo(s) con ffprobe…"

    if ($script:modoTest) {
        # Sin bucle de mensajes en modo test: escaneo síncrono.
        $res = & $scanWorker $rutas
        Aplicar-Analisis @($res)
        return
    }
    $ps = [powershell]::Create()
    [void]$ps.AddScript($scanWorker.ToString()).AddArgument($rutas)
    $script:scanPS = $ps
    $script:scanAsync = $ps.BeginInvoke()
    $script:scanGenLanzado = $script:scanGen
    $script:pollTimer.Start()
}

# Debounce: re-analiza 600 ms después del último cambio en la ruta
$script:debounceTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:debounceTimer.Interval = [TimeSpan]::FromMilliseconds(600)
$script:debounceTimer.Add_Tick({
    $script:debounceTimer.Stop()
    Lanzar-Escaneo
})

# =========================================================================
# EVENTOS
# =========================================================================
$script:todosPaneles = @("panGeneral","panProyecto","panAudio","panSubs","panSubida","panAjustes")
function Mostrar-Panel($nombre) {
    foreach ($p in $script:todosPaneles) { $ui[$p].Visibility = "Collapsed" }
    $ui[$nombre].Visibility = "Visible"
    # Barra inferior: en «Torrent y subida» el botón de acción es «Subir a HDZERO»
    # (y no «Iniciar montaje»); la consola del montaje solo aplica a la rama de montaje.
    $esSubida = ($nombre -eq "panSubida")
    if ($ui.btnSubir)   { $ui.btnSubir.Visibility   = if ($esSubida) { "Visible" } else { "Collapsed" } }
    if ($ui.btnIniciar) { $ui.btnIniciar.Visibility = if ($esSubida) { "Collapsed" } else { "Visible" } }
    if ($ui.btnConsola) { $ui.btnConsola.Visibility = if ($esSubida) { "Collapsed" } else { "Visible" } }
    # Texto contextual de la barra de carpeta: el uso de la carpeta cambia según la pantalla.
    if ($ui.lblCarpetaDesc) {
        $ui.lblCarpetaDesc.Text = switch ($nombre) {
            "panSubida"  { "Carpeta de trabajo: aquí se buscan por defecto los .torrent y las capturas al subir (botón «Último creado» y al localizar el vídeo de un torrent)." }
            "panAjustes" { "Carpeta de trabajo (común a todo el programa). No se usa en Ajustes; configura aquí tus claves." }
            default      { "Carpeta con los vídeos a procesar. Las opciones se adaptan a lo que se detecte; el resultado se monta desde aquí." }
        }
    }
}
$ui.navGeneral.Add_Checked({  Mostrar-Panel "panGeneral" })
$ui.navProyecto.Add_Checked({ Mostrar-Panel "panProyecto" })
$ui.navAudio.Add_Checked({    Mostrar-Panel "panAudio" })
$ui.navSubs.Add_Checked({     Mostrar-Panel "panSubs" })
$ui.navSubida.Add_Checked({
    Mostrar-Panel "panSubida"
    # Al mostrar el panel ya tiene ancho real → recalcular si las pestañas desbordan (flechas).
    $ui.scrTabs.Dispatcher.InvokeAsync({ Update-FlechasTabs }, [System.Windows.Threading.DispatcherPriority]::ContextIdle) | Out-Null
})
$ui.navAjustes.Add_Checked({  Mostrar-Panel "panAjustes" })

# --- Menú lateral colapsable (solo iconos) ---
$script:sidebarColapsado = $false
$script:logoCargado = $false
$script:navItems = @(
    @{ C = $ui.navGeneral;  I = "🏠"; L = "General" },
    @{ C = $ui.navProyecto; I = "🎬"; L = "Proyecto" },
    @{ C = $ui.navAudio;    I = "🔊"; L = "Audio" },
    @{ C = $ui.navSubs;     I = "💬"; L = "Subtítulos" },
    @{ C = $ui.navSubida;   I = "☁"; L = "Torrent y subida" },
    @{ C = $ui.navAjustes;  I = "⚙"; L = "Ajustes / claves" }
)
function Aplicar-Sidebar {
    $col = $script:sidebarColapsado
    $ui.sidebar.Width = if ($col) { 72 } else { 216 }
    $visTxt = if ($col) { "Collapsed" } else { "Visible" }
    foreach ($e in @($ui.lblSecMontaje, $ui.lblSecTorrent, $ui.lblPie, $ui.lblVersion)) { if ($e) { $e.Visibility = $visTxt } }
    $ui.imgLogo.Visibility      = if (-not $col -and $script:logoCargado)      { "Visible" } else { "Collapsed" }
    $ui.panLogoTexto.Visibility = if (-not $col -and -not $script:logoCargado) { "Visible" } else { "Collapsed" }
    $padCol = [System.Windows.Thickness]::new(2, 10, 2, 10)
    $padExp = [System.Windows.Thickness]::new(13, 10, 13, 10)
    foreach ($n in $script:navItems) {
        if ($col) { $n.C.Content = $n.I; $n.C.HorizontalContentAlignment = "Center"; $n.C.FontSize = 18; $n.C.Padding = $padCol; $n.C.ToolTip = $n.L }
        else      { $n.C.Content = "$($n.I)   $($n.L)"; $n.C.HorizontalContentAlignment = "Left"; $n.C.FontSize = 13.5; $n.C.Padding = $padExp; $n.C.ToolTip = $null }
    }
    # Colapsado: el ☰ ocupa las 2 columnas y se centra (la versión está oculta). Expandido: solo
    # su columna derecha, con la versión a la izquierda.
    [System.Windows.Controls.Grid]::SetColumnSpan($ui.btnColapsar, $(if ($col) { 2 } else { 1 }))
    $ui.btnColapsar.HorizontalAlignment = if ($col) { "Center" } else { "Right" }
}
$ui.btnColapsar.Add_Click({ $script:sidebarColapsado = -not $script:sidebarColapsado; Aplicar-Sidebar })

$ui.btnExaminar.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Selecciona la carpeta con los vídeos a procesar"
    if ($ui.txtCarpeta.Text -and (Test-Path -LiteralPath $ui.txtCarpeta.Text)) { $dlg.SelectedPath = $ui.txtCarpeta.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $ui.txtCarpeta.Text = $dlg.SelectedPath }
})

# Quitar el sufijo «.procesado» a los originales de la carpeta para volver a procesarlos.
$ui.btnQuitarProcesado.Add_Click({
    $carpeta = "$($ui.txtCarpeta.Text)".Trim()
    if (-not ($carpeta -and (Test-Path -LiteralPath $carpeta))) { return }
    $procesados = @($script:archivosProcesado)
    if ($procesados.Count -eq 0) { return }
    $ok = Show-DialogoHDZ -Owner $win -Icono "♻️" -Titulo "Quitar «.procesado»" `
        -Mensaje "Se quitará el sufijo «.procesado» de $($procesados.Count) archivo(s), devolviéndoles su nombre original para poder trabajar con ellos de nuevo.`n`n¿Continuar?" `
        -BotonSi "Sí, quitar sufijo" -BotonNo "Cancelar"
    if (-not $ok) { return }
    $restaurados = 0; $omitidos = 0
    foreach ($f in $procesados) {
        $nuevo = $f.Name -replace '\.procesado$', ''
        if ($nuevo -eq $f.Name) { continue }
        $destino = Join-Path $f.DirectoryName $nuevo
        if (Test-Path -LiteralPath $destino) { $omitidos++; continue }   # ya existe el original: no se pisa
        try { Rename-Item -LiteralPath $f.FullName -NewName $nuevo -ErrorAction Stop; $restaurados++ }
        catch { $omitidos++ }
    }
    if ($omitidos -gt 0) {
        Set-Estado "Restaurados $restaurados archivo(s). $omitidos no se pudieron (ya existía un archivo con el nombre original)." "warn"
    } else {
        Set-Estado "Restaurados $restaurados archivo(s): listos para procesar." "ok"
    }
    Lanzar-Escaneo   # re-escanea: los vídeos restaurados ya aparecen en la lista
})
# Carpetas de salida (rama de montaje): selectores de carpeta
$ui.btnSalidaArchivo.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Carpeta donde guardar el MKV procesado y sus capturas"
    if ($ui.txtSalidaArchivo.Text -and (Test-Path -LiteralPath $ui.txtSalidaArchivo.Text)) { $dlg.SelectedPath = $ui.txtSalidaArchivo.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $ui.txtSalidaArchivo.Text = $dlg.SelectedPath }
})
$ui.btnSalidaTorrent.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Carpeta donde guardar el archivo .torrent"
    if ($ui.txtSalidaTorrent.Text -and (Test-Path -LiteralPath $ui.txtSalidaTorrent.Text)) { $dlg.SelectedPath = $ui.txtSalidaTorrent.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $ui.txtSalidaTorrent.Text = $dlg.SelectedPath }
})
$ui.txtCarpeta.Add_TextChanged({
    Actualizar-ListaRapida
    if (-not $script:modoTest) {
        $script:debounceTimer.Stop()
        $script:debounceTimer.Start()
    }
})

$win.Add_PreviewDragOver({
    param($s, $e)
    if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $e.Effects = [Windows.DragDropEffects]::Copy
        $e.Handled = $true
    }
})
$win.Add_PreviewDrop({
    param($s, $e)
    if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $items = $e.Data.GetData([Windows.DataFormats]::FileDrop)
        if ($items -and $items.Count -gt 0) {
            $p = $items[0]
            if (Test-Path -LiteralPath $p -PathType Container) { $ui.txtCarpeta.Text = $p }
            elseif (Test-Path -LiteralPath $p)                 { $ui.txtCarpeta.Text = (Split-Path $p -Parent) }
        }
        $e.Handled = $true
    }
})

$ui.swProyecto.Add_Checked({ Actualizar-Proyecto })
$ui.swProyecto.Add_Unchecked({ Actualizar-Proyecto })

# Selección de archivos: marcar/desmarcar todos
$ui.btnSelTodos.Add_Click({
    foreach ($f in $script:ultimoScan) { $script:seleccion[$f.Nombre] = $true }
    Refrescar-Adaptacion
})
$ui.btnSelNinguno.Add_Click({
    foreach ($f in $script:ultimoScan) { $script:seleccion[$f.Nombre] = $false }
    Refrescar-Adaptacion
})

# =========================================================================
# SUBIR AL TRACKER — catálogos del formulario de HDZERO (IDs reales)
# =========================================================================
$catsHDZ = @(
    @{T="Películas"; V=1; Tipo="movie"}, @{T="Series"; V=2; Tipo="tv"}, @{T="Series Emisión"; V=5; Tipo="tv"},
    @{T="Animación Películas"; V=6; Tipo="movie"}, @{T="Animación Series"; V=7; Tipo="tv"},
    @{T="Anime Películas"; V=8; Tipo="movie"}, @{T="Anime Series"; V=9; Tipo="tv"},
    @{T="Documentales"; V=10; Tipo="movie"}, @{T="Documentales Series"; V=11; Tipo="tv"},
    @{T="Programas TV"; V=12; Tipo="tv"}, @{T="Telenovelas"; V=13; Tipo="tv"}
)
$tiposHDZ = @(
    @{T="WEB-DL"; V=6}, @{T="Full UHD"; V=3}, @{T="Custom UHD"; V=4}, @{T="UHDRemux"; V=5},
    @{T="Full BluRay"; V=1}, @{T="Custom BD"; V=8}, @{T="BDRemux"; V=2}, @{T="Remastered"; V=9},
    @{T="BDrip"; V=7}, @{T="UHDRip"; V=10}, @{T="MicroHD"; V=11}, @{T="DVDRemux"; V=13}, @{T="Rips AV1"; V=14}
)
$resolucionesHDZ = @(
    @{T="2160p"; V=2}, @{T="1080p"; V=3}, @{T="4320p"; V=1}, @{T="1080i"; V=4}, @{T="720p"; V=5},
    @{T="576p"; V=6}, @{T="576i"; V=7}, @{T="480p"; V=8}, @{T="480i"; V=9}, @{T="Other"; V=10}
)
Init-Combo $ui.upCategoria $catsHDZ 0
Init-Combo $ui.upTipo $tiposHDZ 0
Init-Combo $ui.upResolucion $resolucionesHDZ 0
Init-Combo $ui.cmbNumCapturas @(@{T="3";V=3},@{T="4";V=4},@{T="5";V=5},@{T="6";V=6},@{T="8";V=8},@{T="10";V=10}) 3
Add-ChipGroup $ui.chHostImg "hostImg" @(
    @{T="freeimage.host (sin clave)"; V="FREEIMAGE"},
    @{T="imgbb (con clave)"; V="IMGBB"}
) 0 { if ($ui -and $ui.panImgbbKey) { $ui.panImgbbKey.Visibility = if ((Get-ChipValor "hostImg") -eq "IMGBB") { "Visible" } else { "Collapsed" } } }

# Estado de la subida
$script:capRows = @()            # capturas: @{Ruta; Check; Subida(url)}
$script:videoSubir = $null       # ruta del vídeo (o 1er episodio) asociado al .torrent
$script:tmdbTipo = "movie"       # movie|tv según la categoría elegida

# --- Lee el campo 'name' de un .torrent (bencode mínimo) ---
function Get-NombreTorrent($rutaTorrent) {
    try {
        $bytes = [System.IO.File]::ReadAllBytes($rutaTorrent)
        $txt = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
        $m = [regex]::Match($txt, '4:name(\d+):')
        if (-not $m.Success) { return $null }
        $len = [int]$m.Groups[1].Value
        $ini = $m.Index + $m.Length
        $nombreBytes = $bytes[$ini..($ini + $len - 1)]
        return [System.Text.Encoding]::UTF8.GetString($nombreBytes)
    } catch { return $null }
}

# --- Localiza el vídeo (o 1er episodio si es carpeta/pack) a partir del nombre del torrent ---
function Resolve-VideoDeTorrent($rutaTorrent, $nombreInterno) {
    $dir = Split-Path $rutaTorrent -Parent
    $cand = Join-Path $dir $nombreInterno
    if (Test-Path -LiteralPath $cand -PathType Leaf) { return $cand }            # single-file
    if (Test-Path -LiteralPath $cand -PathType Container) {                       # pack: 1er vídeo
        $v = @(Get-ChildItem -LiteralPath $cand -File -Recurse |
               Where-Object { $_.Extension -match "(?i)^\.(mkv|mp4|ts)$" } | Sort-Object FullName | Select-Object -First 1)
        if ($v.Count -gt 0) { return $v[0].FullName }
    }
    # Respaldo: por nombre base en la carpeta del .torrent
    $base = [System.IO.Path]::GetFileNameWithoutExtension($nombreInterno)
    $v2 = @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match "(?i)^\.(mkv|mp4|ts)$" -and $_.BaseName -eq $base } | Select-Object -First 1)
    if ($v2.Count -gt 0) { return $v2[0].FullName }
    # Respaldo extra: el vídeo puede estar en una carpeta de salida distinta a la del .torrent.
    # Buscamos por nombre base (recursivo) en las carpetas conocidas: salida del archivo y trabajo.
    $otras = @("$($ui.txtSalidaArchivo.Text)", "$($ui.txtCarpeta.Text)") |
             Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
    foreach ($od in $otras) {
        if ($od -eq $dir) { continue }
        $vx = @(Get-ChildItem -LiteralPath $od -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match "(?i)^\.(mkv|mp4|ts)$" -and $_.BaseName -eq $base } | Select-Object -First 1)
        if ($vx.Count -gt 0) { return $vx[0].FullName }
    }
    return $null
}

# --- MediaInfo en texto plano completo ---
function Get-MediaInfoTexto($ruta) {
    try { return (& mediainfo "$ruta" | Out-String).TrimEnd() } catch { return "" }
}

# --- Detección de specs desde mediainfo JSON (para preseleccionar combos y flags) ---
function Get-SpecsVideo($ruta) {
    $r = @{ Resolucion=$null; EsWebDl=$false; DV=$false; HDR10=$false; HDR10P=$false; Codec="" }
    try {
        $j = & mediainfo --Output=JSON "$ruta" | Out-String | ConvertFrom-Json
        $v = $j.media.track | Where-Object { $_.'@type' -eq "Video" } | Select-Object -First 1
        if ($v) {
            $alto = [int]"0$($v.Height)"
            $r.Resolucion = if ($alto -gt 1200) { 2 } elseif ($alto -gt 700) { 3 } elseif ($alto -gt 0) { 5 } else { $null }
            $hdr = "$($v.HDR_Format) $($v.HDR_Format_Compatibility) $($v.HDR_Format_String)"
            if ($hdr -match "(?i)Dolby Vision") { $r.DV = $true }
            if ($hdr -match "(?i)HDR10\+|SMPTE ST 2094") { $r.HDR10P = $true }
            if ($hdr -match "(?i)HDR10") { $r.HDR10 = $true }
            $r.Codec = "$($v.Format)"
        }
    } catch {}
    return $r
}

# --- Sube una imagen al host elegido; devuelve la URL full-size o $null ---
function Send-Imagen($ruta) {
    $hostImg = Get-ChipValor "hostImg"
    try {
        if ($hostImg -eq "IMGBB") {
            $key = "$($ui.cfgImgbbKey.Text)".Trim()
            if ([string]::IsNullOrWhiteSpace($key)) { throw "falta la clave de imgbb (pestaña Ajustes)" }
            $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($ruta))
            $resp = Invoke-RestMethod -Uri "https://api.imgbb.com/1/upload?key=$key" -Method Post -Body @{ image = $b64 } -TimeoutSec 120
            if ($resp.success) { return "$($resp.data.url)" }
            throw "imgbb no devolvió URL"
        } else {
            $form = @{ key = "6d207e02198a847aa98d0a2a901485a5"; action = "upload"; format = "json"; source = Get-Item -LiteralPath $ruta }
            $resp = Invoke-RestMethod -Uri "https://freeimage.host/api/1/upload" -Method Post -Form $form -TimeoutSec 120
            if ("$($resp.status_code)" -eq "200") { return "$($resp.image.url)" }
            throw "freeimage devolvió estado $($resp.status_code)"
        }
    } catch { throw }
}

# Worker que sube una lista de imágenes EN UN RUNSPACE APARTE (no bloquea la UI). Comparte el
# progreso por un hashtable sincronizado ($sync). Replica la lógica de Send-Imagen sin tocar $ui.
$subidaWorker = {
    param($rutas, $hostType, $imgbbKey, $sync)
    $urls = @()
    foreach ($ruta in @($rutas)) {
        try {
            if ($hostType -eq "IMGBB") {
                $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($ruta))
                $resp = Invoke-RestMethod -Uri "https://api.imgbb.com/1/upload?key=$imgbbKey" -Method Post -Body @{ image = $b64 } -TimeoutSec 120
                if ($resp.success) { $urls += "$($resp.data.url)" } else { throw "imgbb no devolvió URL" }
            } else {
                $form = @{ key = "6d207e02198a847aa98d0a2a901485a5"; action = "upload"; format = "json"; source = Get-Item -LiteralPath $ruta }
                $resp = Invoke-RestMethod -Uri "https://freeimage.host/api/1/upload" -Method Post -Form $form -TimeoutSec 120
                if ("$($resp.status_code)" -eq "200") { $urls += "$($resp.image.url)" } else { throw "freeimage devolvió estado $($resp.status_code)" }
            }
        } catch {
            $sync.Error = "Error subiendo $([System.IO.Path]::GetFileName($ruta)): $($_.Exception.Message)"
            $sync.Urls = $urls; $sync.Fin = $true; return
        }
        $sync.Done = $sync.Done + 1
    }
    $sync.Urls = $urls; $sync.Fin = $true
}

# Sube imágenes de forma ASÍNCRONA con barra de progreso; al terminar llama a $onComplete(urls)
# en el hilo de UI. $botones se deshabilitan mientras dura. La UI NO se bloquea.
$script:subTimer = $null; $script:subSync = $null; $script:subPS = $null
$script:subAsync = $null; $script:subOnComplete = $null; $script:subBotones = @()
function Start-SubidaAsync($rutas, [scriptblock]$onComplete, $botones = @()) {
    $rutas = @($rutas)
    if ($rutas.Count -eq 0) { return }
    $hostType = Get-ChipValor "hostImg"
    $imgbbKey = "$($ui.cfgImgbbKey.Text)".Trim()
    if ($hostType -eq "IMGBB" -and [string]::IsNullOrWhiteSpace($imgbbKey)) {
        Set-Estado "Falta la clave de imgbb (pestaña Ajustes)." "error"; return
    }
    $script:subOnComplete = $onComplete
    $script:subBotones = @($botones)
    foreach ($b in $script:subBotones) { if ($b) { $b.IsEnabled = $false } }

    $sync = [hashtable]::Synchronized(@{ Done = 0; Total = $rutas.Count; Urls = @(); Fin = $false; Error = $null })
    $script:subSync = $sync
    $ui.panSubProgreso.Visibility = "Visible"
    $ui.barSubProgreso.Value = 0
    $ui.lblSubProgreso.Text = "Subiendo imágenes al host…"
    $ui.lblSubProgresoPct.Text = "0/$($rutas.Count)"

    $ps = [powershell]::Create()
    [void]$ps.AddScript($subidaWorker.ToString()).AddArgument($rutas).AddArgument($hostType).AddArgument($imgbbKey).AddArgument($sync)
    $script:subPS = $ps
    $script:subAsync = $ps.BeginInvoke()

    if (-not $script:subTimer) {
        $script:subTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:subTimer.Interval = [TimeSpan]::FromMilliseconds(200)
        $script:subTimer.Add_Tick({
            $s = $script:subSync
            if (-not $s) { $script:subTimer.Stop(); return }
            $done = [int]$s.Done; $tot = [int]$s.Total
            $ui.barSubProgreso.Value = if ($tot -gt 0) { [int](100 * $done / $tot) } else { 0 }
            $ui.lblSubProgresoPct.Text = "$done/$tot"
            if ($s.Fin) {
                $script:subTimer.Stop()
                try { $script:subPS.EndInvoke($script:subAsync) } catch {}
                try { $script:subPS.Dispose() } catch {}
                $script:subPS = $null; $script:subAsync = $null
                $ui.panSubProgreso.Visibility = "Collapsed"
                foreach ($b in $script:subBotones) { if ($b) { $b.IsEnabled = $true } }
                $cb = $script:subOnComplete; $err = $s.Error; $urls = @($s.Urls)
                $script:subSync = $null; $script:subOnComplete = $null
                if ($err) { Set-Estado $err "error" }
                elseif ($cb) { & $cb $urls }
            }
        })
    }
    $script:subTimer.Start()
}

# Genera capturas en el momento (mismo cálculo y ffmpeg que el motor New-Capturas de HDZnew),
# en un runspace aparte para no bloquear la UI. Reporta progreso por hashtable sincronizado y
# escribe los JPG "<base>_cap_<seg>.jpg" junto al vídeo. $sync.Hechas = rutas creadas.
$capturaWorker = {
    param($video, $numCapturas, $esHdr, $sync)
    try {
        $rutaCarpeta = Split-Path $video -Parent
        $base = [System.IO.Path]::GetFileNameWithoutExtension($video)
        $rawDur = "$(ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "$video" 2>$null)".Trim()
        $dur = 0.0
        $ok = (-not [string]::IsNullOrWhiteSpace($rawDur)) -and
              [double]::TryParse($rawDur.Replace(',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dur) -and
              $dur -gt 0
        if (-not $ok) { $sync.Error = "No se pudo leer la duración del vídeo con ffprobe."; $sync.Fin = $true; return }

        if ($dur -le 300)       { $inic = 20 }
        elseif ($dur -ge 1800)  { $inic = 300 }
        else                    { $inic = 20 + (300 - 20) * ($dur - 300) / (1800 - 300) }
        $margenFin = if ($dur -lt 1800) { [Math]::Min(60, $dur * 0.08) } else { 60 }
        $fin = $dur - $margenFin
        if ($fin -le $inic) { $inic = $dur * 0.10; $fin = $dur * 0.90 }
        $ventana = $fin - $inic

        $creadas = @()
        for ($i = 0; $i -lt $numCapturas; $i++) {
            $t = if ($numCapturas -eq 1) { [int]($inic + $ventana / 2) } else { [int]($inic + ($ventana * $i / ($numCapturas - 1))) }
            $ts = $t.ToString("D4")
            $rutaCap = Join-Path $rutaCarpeta "${base}_cap_$ts.jpg"
            if ($esHdr) {
                ffmpeg -noaccurate_seek -ss $t -i "$video" -frames:v 1 -vf "zscale=t=linear:npl=250,tonemap=tonemap=reinhard:desat=2,zscale=p=709:t=709:m=709,format=yuv420p" -q:v 2 "$rutaCap" -y -loglevel fatal
            } else {
                ffmpeg -noaccurate_seek -ss $t -i "$video" -frames:v 1 -q:v 2 "$rutaCap" -y -loglevel fatal
            }
            if (Test-Path -LiteralPath $rutaCap) { $creadas += $rutaCap }
            $sync.Done = $sync.Done + 1
        }
        $sync.Hechas = $creadas
    } catch { $sync.Error = $_.Exception.Message }
    $sync.Fin = $true
}

# Crea capturas de forma ASÍNCRONA con barra de progreso (reutiliza panSubProgreso). Al terminar
# refresca la rejilla de capturas con las antiguas + las nuevas.
$script:capTimer = $null; $script:capSync = $null; $script:capPS = $null; $script:capAsync = $null
function Start-CapturasAsync {
    if (-not ($script:videoSubir -and (Test-Path -LiteralPath $script:videoSubir))) {
        Set-Estado "No hay vídeo asociado: carga o crea un torrent primero para saber de qué vídeo sacar las capturas." "error"; return
    }
    $num = Get-ComboValor $ui.cmbNumCapturas; if (-not $num) { $num = 6 }
    $esHdr = [bool]$ui.upDV.IsChecked -or [bool]$ui.upHDR10.IsChecked -or [bool]$ui.upHDR10P.IsChecked
    $sync = [hashtable]::Synchronized(@{ Done = 0; Total = [int]$num; Hechas = @(); Fin = $false; Error = $null })
    $script:capSync = $sync
    foreach ($b in @($ui.btnCapGenerar, $ui.btnCapInsertar, $ui.btnCapAnadir)) { if ($b) { $b.IsEnabled = $false } }
    $ui.panSubProgreso.Visibility = "Visible"
    $ui.barSubProgreso.Value = 0
    $ui.lblSubProgreso.Text = "Generando capturas$(if ($esHdr) { ' (tonemap HDR)' } else { '' })…"
    $ui.lblSubProgresoPct.Text = "0/$num"

    $ps = [powershell]::Create()
    [void]$ps.AddScript($capturaWorker.ToString()).AddArgument($script:videoSubir).AddArgument([int]$num).AddArgument($esHdr).AddArgument($sync)
    $script:capPS = $ps
    $script:capAsync = $ps.BeginInvoke()

    if (-not $script:capTimer) {
        $script:capTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:capTimer.Interval = [TimeSpan]::FromMilliseconds(200)
        $script:capTimer.Add_Tick({
            $s = $script:capSync
            if (-not $s) { $script:capTimer.Stop(); return }
            $done = [int]$s.Done; $tot = [int]$s.Total
            $ui.barSubProgreso.Value = if ($tot -gt 0) { [int](100 * $done / $tot) } else { 0 }
            $ui.lblSubProgresoPct.Text = "$done/$tot"
            if ($s.Fin) {
                $script:capTimer.Stop()
                try { $script:capPS.EndInvoke($script:capAsync) } catch {}
                try { $script:capPS.Dispose() } catch {}
                $script:capPS = $null; $script:capAsync = $null
                $ui.panSubProgreso.Visibility = "Collapsed"
                foreach ($b in @($ui.btnCapGenerar, $ui.btnCapInsertar, $ui.btnCapAnadir)) { if ($b) { $b.IsEnabled = $true } }
                $err = $s.Error; $nuevas = @($s.Hechas); $script:capSync = $null
                if ($err) { Set-Estado "No se pudieron crear las capturas: $err" "error"; return }
                # Une las que ya había con las nuevas (sin duplicar) y refresca la rejilla.
                $previas = @($script:capRows | ForEach-Object { $_.Ruta })
                $todas = @($previas + $nuevas | Select-Object -Unique)
                Construir-Capturas $todas
                Set-Estado "✓  $($nuevas.Count) captura(s) creada(s) junto al vídeo." "ok"
            }
        })
    }
    $script:capTimer.Start()
}

# --- Llamada genérica a TMDB. Acepta INDISTINTAMENTE las dos credenciales de TMDB:
#     - "Clave de la API" (v3, 32 hex)         -> va como ?api_key=
#     - "Token de acceso de lectura" (v4, JWT) -> va como header Authorization: Bearer
#     Se detecta por el formato ("eyJ..." = JWT). Así funciona pegue el usuario la que pegue. ---
function Invoke-Tmdb($ruta, $params = @{}) {
    $cred = "$($ui.cfgTmdbKey.Text)".Trim()
    if ([string]::IsNullOrWhiteSpace($cred)) { throw "falta la clave/token de TMDB (pestaña Ajustes)" }
    $esJwt = ($cred -match '^eyJ' -or $cred.Length -gt 60)
    $qs = @()
    foreach ($k in $params.Keys) { if ("$($params[$k])" -ne "") { $qs += "$k=$([uri]::EscapeDataString("$($params[$k])"))" } }
    if (-not $esJwt) { $qs = @("api_key=$cred") + $qs }
    $uri = "https://api.themoviedb.org/3/$ruta" + $(if ($qs.Count) { "?" + ($qs -join "&") } else { "" })
    $headers = @{ Accept = "application/json" }
    if ($esJwt) { $headers["Authorization"] = "Bearer $cred" }
    return Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -TimeoutSec 60
}

# --- Búsqueda TMDB (movie|tv) tolerante con el año, como UNIT3D ---
# El año NUNCA excluye resultados: se busca sin año (base amplia) y, si se indicó,
# también con año para priorizar; luego se ordena por coincidencia/cercanía de año y
# popularidad. Así, aunque el año del nombre esté mal, el título sigue encontrando.
function Search-Tmdb($titulo, $anio, $tipo) {
    $campoAnio = if ($tipo -eq "tv") { "first_air_date_year" } else { "year" }
    $vistos = @{}
    $acc = New-Object System.Collections.Generic.List[object]
    $agregar = {
        param($res, $conAnio)
        foreach ($x in @($res)) {
            if (-not $x -or $vistos.ContainsKey("$($x.id)")) { continue }
            $vistos["$($x.id)"] = $true
            $fecha = if ($tipo -eq "tv") { "$($x.first_air_date)" } else { "$($x.release_date)" }
            $aR = if ($fecha -match "^(\d{4})") { [int]$Matches[1] } else { 0 }
            $acc.Add([PSCustomObject]@{
                Id = $x.id
                Titulo = if ($tipo -eq "tv") { $x.name } else { $x.title }
                Fecha = $fecha
                AnioR = $aR
                Pop = [double]("0$($x.popularity)" -replace ',', '.')
                CoincideAnio = ($conAnio -and $aR -gt 0 -and $anio -and [int]$anio -eq $aR)
                Generos = @($x.genre_ids)
                Lang = "$($x.original_language)"
            })
        }
    }
    # 1) con año (los resultados más precisos, si el año es correcto)
    if ($anio) {
        try { & $agregar (Invoke-Tmdb "search/$tipo" @{ language="es-ES"; query=$titulo; $campoAnio=$anio }).results $true } catch {}
    }
    # 2) sin año SIEMPRE: aunque el año del nombre esté mal, el título sigue encontrando
    try { & $agregar (Invoke-Tmdb "search/$tipo" @{ language="es-ES"; query=$titulo }).results $false } catch {}

    # Orden: coincidencia exacta de año > cercanía al año pedido > popularidad
    $aPedido = if ($anio) { [int]$anio } else { 0 }
    return @($acc | Sort-Object `
        @{ Expression = { if ($_.CoincideAnio) { 0 } else { 1 } } }, `
        @{ Expression = { if ($aPedido -and $_.AnioR) { [Math]::Abs($_.AnioR - $aPedido) } else { 9999 } } }, `
        @{ Expression = { $_.Pop }; Descending = $true } | Select-Object -First 10)
}

# --- IDs externos (imdb, tvdb) de un elemento de TMDB ---
function Get-TmdbExternos($id, $tipo) {
    try { return Invoke-Tmdb "$tipo/$id/external_ids" }
    catch { return $null }
}

# --- Palabras clave de TMDB (string separado por comas), como hace la web ---
function Get-TmdbKeywords($id, $tipo) {
    try {
        $r = Invoke-Tmdb "$tipo/$id/keywords"
        $lista = if ($tipo -eq "tv") { $r.results } else { $r.keywords }
        return (@($lista | ForEach-Object { $_.name } | Where-Object { $_ }) -join ", ")
    } catch { return "" }
}

# --- Logos de TMDB (full-size + dimensiones); prioriza español, luego inglés, luego sin idioma ---
function Get-TmdbLogos($id, $tipo) {
    try {
        $r = Invoke-Tmdb "$tipo/$id/images" @{ include_image_language = "es,en,null" }
        $todos = @($r.logos | ForEach-Object {
            [PSCustomObject]@{
                Url = "https://image.tmdb.org/t/p/original$($_.file_path)"
                Width = [int]"0$($_.width)"; Height = [int]"0$($_.height)"
                Lang = "$($_.iso_639_1)"
                Voto = [double]("0$($_.vote_average)" -replace ',', '.')
            }
        })
        # Mostramos TODOS, pero ordenados: primero los de español (es / es-419), luego inglés, luego
        # el resto/sin idioma; dentro de cada grupo, por valoración descendente.
        $rango = { param($l) if ($l.Lang -eq "es" -or $l.Lang -eq "es-419") { 0 } elseif ($l.Lang -eq "en") { 1 } else { 2 } }
        return @($todos | Sort-Object `
            @{ Expression = { & $rango $_ } }, `
            @{ Expression = { $_.Voto }; Descending = $true })
    } catch { return @() }
}

# --- Limpia un nombre de release para buscar en TMDB: devuelve @{Titulo; Anio} ---
function Get-TituloAnioLimpio($titulo) {
    $t = "$titulo"
    $anio = ""
    if ($t -match "\b(19\d{2}|20\d{2})\b") { $anio = $Matches[1] }
    $t = ($t -replace "(?i)\bS\d{1,3}(E\d{1,4})?\b.*$", "")     # corta en SxxEyy
    $t = ($t -replace "\(?\b(19\d{2}|20\d{2})\b\)?.*$", "")      # corta en el año
    $t = ($t -replace "[._]", " ").Trim()
    return @{ Titulo = $t; Anio = $anio }
}

# =========================================================================
# SUBIR AL TRACKER — lógica de interfaz
# =========================================================================
function Actualizar-TipoCategoria {
    $cat = Get-ComboValor $ui.upCategoria
    $entry = $catsHDZ | Where-Object { $_.V -eq $cat } | Select-Object -First 1
    $script:tmdbTipo = if ($entry -and $entry.Tipo -eq "tv") { "tv" } else { "movie" }
    $ui.panTV.Visibility = if ($script:tmdbTipo -eq "tv") { "Visible" } else { "Collapsed" }
}

# Plantilla de descripción centrada (la descripción siempre va dentro de [center]…[/center])
$script:descVacia = "[center]`n`n[/center]"
function Reset-Descripcion { $ui.upDescripcion.Text = $script:descVacia }

# Inserta/envuelve la selección de la descripción con etiquetas BBCode
function Insert-Bb($apertura, $cierre) {
    $tb = $ui.upDescripcion
    $s = $tb.SelectionStart; $l = $tb.SelectionLength
    if ($s -lt 0) { $s = $tb.Text.Length; $l = 0 }
    $sel = $tb.Text.Substring($s, $l)
    $rep = "$apertura$sel$cierre"
    $tb.Text = $tb.Text.Remove($s, $l).Insert($s, $rep)
    $tb.SelectionStart = $s + $apertura.Length
    $tb.SelectionLength = $sel.Length
    $tb.Focus()
}

# Pone el logo como primera línea dentro de [center] (sustituye uno previo si lo había)
function Insert-LogoEnDescripcion($url) {
    $tb = $ui.upDescripcion
    $txt = $tb.Text
    if ([string]::IsNullOrWhiteSpace($txt)) { $txt = $script:descVacia }
    $txt = [regex]::Replace($txt, '(?i)\[img=700\][^\[\]]*\[/img\]\r?\n?', '')   # quita logo anterior
    $linea = "[img=700]$url" + "[/img]"
    $idx = $txt.IndexOf("[center]")
    if ($idx -ge 0) {
        $pos = $idx + "[center]".Length
        $tb.Text = $txt.Substring(0, $pos) + "`n" + $linea + $txt.Substring($pos)
    } else {
        $tb.Text = "[center]`n$linea`n`n$txt`n[/center]"
    }
}

# Inserta imágenes [img=350] antes del [/center] de cierre.
# Formato: 5 capturas por línea, separadas por 3 espacios; cada bloque de 5 en una línea nueva.
function Insert-ImagenesEnDescripcion($urls) {
    if (-not $urls -or @($urls).Count -eq 0) { return }
    $tb = $ui.upDescripcion
    $txt = $tb.Text
    if ([string]::IsNullOrWhiteSpace($txt)) { $txt = $script:descVacia }
    $tags = @($urls | ForEach-Object { "[img=350]$_" + "[/img]" })
    $lineas = @()
    for ($i = 0; $i -lt $tags.Count; $i += 5) {
        $fin = [Math]::Min($i + 4, $tags.Count - 1)
        $lineas += ($tags[$i..$fin] -join "   ")   # 3 espacios entre capturas
    }
    $bloque = $lineas -join "`n"
    $idx = $txt.LastIndexOf("[/center]")
    if ($idx -ge 0) {
        $tb.Text = $txt.Substring(0, $idx).TrimEnd() + "`n" + $bloque + "`n" + $txt.Substring($idx)
    } else {
        $tb.Text = $txt.TrimEnd() + "`n[center]$bloque[/center]"
    }
}

# Resalta el logo actualmente elegido (el que se insertó en la descripción)
$script:logoSel = $null
$script:logoBordes = @()
function Refresh-LogoSel {
    foreach ($e in $script:logoBordes) {
        if ($e.Url -eq $script:logoSel) {
            $e.Borde.BorderBrush = $win.Resources["AccentBrush"]; $e.Borde.BorderThickness = [System.Windows.Thickness]::new(2)
        } else {
            $e.Borde.BorderBrush = (Brocha "#33333A"); $e.Borde.BorderThickness = [System.Windows.Thickness]::new(1)
        }
    }
}

# Muestra los logos de TMDB como miniaturas clicables, con su tamaño debajo
function Construir-Logos($logos) {
    $ui.panLogos.Children.Clear()
    $script:logoBordes = @()
    $script:logoSel = $null
    $script:logosActuales = @($logos)   # guardado para snapshot/restore de pestañas
    if (-not $logos -or @($logos).Count -eq 0) {
        $ui.cardLogos.Visibility = "Collapsed"
        return
    }
    $ui.cardLogos.Visibility = "Visible"
    $ui.bdgLogos.Text = "$(@($logos).Count) logo(s)"
    foreach ($lg in @($logos | Select-Object -First 12)) {
        $cont = New-Object System.Windows.Controls.StackPanel
        $cont.Margin = [System.Windows.Thickness]::new(0, 0, 8, 8)
        $b = New-Object System.Windows.Controls.Border
        $b.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $b.Background = Brocha "#3A3A42"      # gris medio: hace visibles logos blancos y oscuros
        $b.BorderBrush = Brocha "#33333A"; $b.BorderThickness = [System.Windows.Thickness]::new(1)
        $b.Padding = [System.Windows.Thickness]::new(8, 6, 8, 6)
        $b.Cursor = [System.Windows.Input.Cursors]::Hand
        $b.ToolTip = "Usar este logo  ·  $($lg.Width)×$($lg.Height) px"
        try {
            $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
            $bmp.BeginInit(); $bmp.UriSource = [Uri]$lg.Url; $bmp.DecodePixelWidth = 240
            $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad; $bmp.EndInit()
            $img = New-Object System.Windows.Controls.Image
            $img.Source = $bmp; $img.Stretch = "Uniform"; $img.Height = 42; $img.MaxWidth = 240
            $b.Child = $img
        } catch {
            $t = New-Object System.Windows.Controls.TextBlock; $t.Text = "logo"; $t.Foreground = $win.Resources["SubBrush"]
            $b.Child = $t
        }
        [void]$cont.Children.Add($b)
        $tam = New-Object System.Windows.Controls.TextBlock
        $tam.Text = "$($lg.Width)×$($lg.Height)$(if ($lg.Lang) { "  ·  $($lg.Lang)" })"
        $tam.FontSize = 10.5; $tam.Foreground = $win.Resources["SubBrush"]
        $tam.HorizontalAlignment = "Center"; $tam.Margin = [System.Windows.Thickness]::new(0, 3, 0, 0)
        [void]$cont.Children.Add($tam)
        # OJO: handler SIN .GetNewClosure() (igual que en las firmas): así la escritura de
        # $script:logoSel propaga al scope real y el resaltado funciona. La ruta va en el Tag.
        $b.Tag = $lg.Url
        $b.Add_MouseLeftButtonUp({
            param($sender, $e)
            $script:logoSel = $sender.Tag
            Insert-LogoEnDescripcion $sender.Tag
            Refresh-LogoSel
            Set-Estado "Logo añadido a la descripción (resaltado abajo)."
        })
        $script:logoBordes += [PSCustomObject]@{ Url = $lg.Url; Borde = $b }
        [void]$ui.panLogos.Children.Add($cont)
    }
}

# --- Firmas: imágenes de la carpeta «firmas» junto al script ---
$script:firmaSel = $null          # ruta de la firma elegida (o $null)
$script:firmaBordes = @()         # @{Ruta;Borde} para resaltar la selección
$script:firmaCacheUrl = @{}       # ruta -> URL subida (cache de sesión)
$rutaFirmas = Join-Path $PSScriptRoot "firmas"

function Refresh-FirmaSel {
    foreach ($e in $script:firmaBordes) {
        if ($e.Ruta -eq $script:firmaSel) { $e.Borde.BorderBrush = $win.Resources["AccentBrush"]; $e.Borde.BorderThickness = [System.Windows.Thickness]::new(2) }
        else { $e.Borde.BorderBrush = (Brocha "#33333A"); $e.Borde.BorderThickness = [System.Windows.Thickness]::new(1) }
    }
}
function Construir-Firmas {
    $ui.panFirmas.Children.Clear()
    $script:firmaBordes = @()
    if (-not (Test-Path -LiteralPath $rutaFirmas)) { $ui.cardFirmas.Visibility = "Collapsed"; return }
    $imgs = @(Get-ChildItem -LiteralPath $rutaFirmas -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Extension -match "(?i)^\.(jpg|jpeg|png|gif|webp)$" } | Sort-Object Name)
    if ($imgs.Count -eq 0) { $ui.cardFirmas.Visibility = "Collapsed"; return }
    $ui.cardFirmas.Visibility = "Visible"
    $ui.bdgFirmas.Text = "$($imgs.Count) firma(s) en la carpeta"
    foreach ($f in $imgs) {
        $cont = New-Object System.Windows.Controls.StackPanel
        $cont.Width = 220; $cont.Margin = [System.Windows.Thickness]::new(0, 0, 10, 10)
        $b = New-Object System.Windows.Controls.Border
        $b.CornerRadius = [System.Windows.CornerRadius]::new(6); $b.Background = Brocha "#14141E"
        $b.BorderBrush = Brocha "#33333A"; $b.BorderThickness = [System.Windows.Thickness]::new(1)
        $b.Padding = [System.Windows.Thickness]::new(6); $b.Cursor = [System.Windows.Input.Cursors]::Hand
        try {
            $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
            $bmp.BeginInit(); $bmp.UriSource = [Uri]$f.FullName; $bmp.DecodePixelWidth = 220
            $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad; $bmp.EndInit()
            $img = New-Object System.Windows.Controls.Image
            $img.Source = $bmp; $img.Stretch = "Uniform"; $img.MaxHeight = 70
            $b.Child = $img
        } catch {}
        [void]$cont.Children.Add($b)
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $f.Name; $t.FontSize = 10.5; $t.Foreground = $win.Resources["SubBrush"]
        $t.TextTrimming = "CharacterEllipsis"; $t.MaxWidth = 210; $t.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
        [void]$cont.Children.Add($t)
        # OJO: nada de .GetNewClosure() aquí. Con GetNewClosure, "$script:firmaSel = …" escribe en
        # una copia aislada del scope y la selección no se propaga (bug que dejaba la firma sin
        # seleccionar). Usamos un handler normal y leemos la ruta del Tag del propio borde.
        $b.Tag = $f.FullName
        $b.Add_MouseLeftButtonUp({
            param($sender, $e)
            $ruta = $sender.Tag
            if ($script:firmaSel -eq $ruta) { $script:firmaSel = $null } else { $script:firmaSel = $ruta }
            Refresh-FirmaSel
        })
        $script:firmaBordes += [PSCustomObject]@{ Ruta = $f.FullName; Borde = $b }
        [void]$ui.panFirmas.Children.Add($cont)
    }
    Refresh-FirmaSel
}

# Aplica un resultado de TMDB: IDs externos, keywords, logos y categoría de animación
function Aplicar-ResultadoTmdb($item) {
    $ui.upTmdb.Text = "$($item.Id)"
    # Animación → categoría correcta (Películas→Animación Películas, Series→Animación Series).
    # Solo si la categoría está en la base (1/2): no piso una elección manual de Anime/otra.
    if (@($item.Generos) -contains 16) {
        $cat = Get-ComboValor $ui.upCategoria
        if ($script:tmdbTipo -eq "movie" -and $cat -eq 1) { Set-ComboValor $ui.upCategoria 6 }
        elseif ($script:tmdbTipo -eq "tv" -and $cat -eq 2) { Set-ComboValor $ui.upCategoria 7 }
    }
    $ext = Get-TmdbExternos $item.Id $script:tmdbTipo
    if ($ext) {
        if ($ext.imdb_id) { $ui.upImdb.Text = "$($ext.imdb_id)" }
        if ($ext.tvdb_id) { $ui.upTvdb.Text = "$($ext.tvdb_id)" }
    }
    $kw = Get-TmdbKeywords $item.Id $script:tmdbTipo
    if ($kw) { $ui.upKeywords.Text = $kw }
    Construir-Logos (Get-TmdbLogos $item.Id $script:tmdbTipo)
}

# Construye las miniaturas de capturas (con check de selección y clic = abrir full-size)
function Construir-Capturas($rutas) {
    $ui.panCapturas.Children.Clear()
    $script:capRows = @()
    foreach ($rc in $rutas) {
        $cont = New-Object System.Windows.Controls.StackPanel
        $cont.Width = 380
        $cont.Margin = [System.Windows.Thickness]::new(0, 0, 12, 12)
        $bordeImg = New-Object System.Windows.Controls.Border
        $bordeImg.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $bordeImg.Background = Brocha "#14141E"
        $bordeImg.Height = 214                 # 380×214 ≈ 16:9, miniatura grande para ver el detalle
        $bordeImg.Cursor = [System.Windows.Input.Cursors]::Hand
        $bordeImg.ClipToBounds = $true
        $bordeImg.ToolTip = "Clic para abrir a tamaño completo"
        try {
            $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
            $bmp.BeginInit(); $bmp.UriSource = [Uri]$rc; $bmp.DecodePixelWidth = 760   # 2× para nitidez
            $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad; $bmp.EndInit()
            $img = New-Object System.Windows.Controls.Image
            $img.Source = $bmp; $img.Stretch = "UniformToFill"
            $bordeImg.Child = $img
        } catch {}
        $rutaCap = $rc
        $bordeImg.Add_MouseLeftButtonUp({ Start-Process $rutaCap }.GetNewClosure())
        [void]$cont.Children.Add($bordeImg)
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Style = $win.Resources["Check"]
        $cb.IsChecked = $true
        $cb.Margin = [System.Windows.Thickness]::new(0, 6, 0, 0)
        $tbn = New-Object System.Windows.Controls.TextBlock
        $tbn.Text = Split-Path $rc -Leaf
        $tbn.FontSize = 10.5; $tbn.Foreground = $win.Resources["SubBrush"]
        $tbn.TextTrimming = "CharacterEllipsis"; $tbn.MaxWidth = 350
        $cb.Content = $tbn
        [void]$cont.Children.Add($cb)
        [void]$ui.panCapturas.Children.Add($cont)
        $script:capRows += [PSCustomObject]@{ Ruta = $rc; Check = $cb; Subida = $null }
    }
    $ui.bdgCapturas.Text = "$($rutas.Count) captura(s) encontradas"
}

# Carga un .torrent: localiza vídeo, mediainfo, specs, capturas, prerellena campos
function Cargar-TorrentParaSubir($rutaTorrent, $videoHint = $null, $origenHint = "", $capturasHint = @()) {
    $ui.txtTorrentSubir.Text = $rutaTorrent
    $nombre = Get-NombreTorrent $rutaTorrent
    if (-not $nombre) {
        $ui.lblTorrentInfo.Text = "No se pudo leer el nombre interno del .torrent."
        $ui.lblTorrentInfo.Visibility = "Visible"
        return
    }
    $ui.lblTorrentInfo.Visibility = "Collapsed"
    # Cada torrent arranca con descripción centrada limpia y sin logos del anterior
    Reset-Descripcion
    Construir-Logos @()
    # Título: nombre interno sin extensión
    $ui.upTitulo.Text = [System.IO.Path]::GetFileNameWithoutExtension($nombre)

    # Si nos dan el vídeo (al crear el torrent ya sabemos cuál es), lo usamos; si no, lo localizamos.
    $video = if ($videoHint -and (Test-Path -LiteralPath $videoHint)) { $videoHint } else { Resolve-VideoDeTorrent $rutaTorrent $nombre }
    $script:videoSubir = $video
    if (-not $video) {
        $ui.lblTorrentInfo.Text = "Torrent cargado, pero no encontré el vídeo junto al .torrent. El MediaInfo y las specs habrá que ponerlos a mano."
        $ui.lblTorrentInfo.Visibility = "Visible"
    } else {
        Set-Estado "Extrayendo MediaInfo y specs de $(Split-Path $video -Leaf)…"
        $ui.upMediainfo.Text = Get-MediaInfoTexto $video
        $specs = Get-SpecsVideo $video
        if ($specs.Resolucion) { Set-ComboValor $ui.upResolucion $specs.Resolucion }
        $ui.upDV.IsChecked = $specs.DV
        $ui.upHDR10P.IsChecked = $specs.HDR10P
        $ui.upHDR10.IsChecked = $specs.HDR10
    }

    # Categoría/Tipo por el nombre (heurística): serie si hay SxxEyy o Sxx
    $esSerie = ($nombre -match "(?i)\bS\d{1,3}(E\d{1,4})?\b")
    Set-ComboValor $ui.upCategoria $(if ($esSerie) { 2 } else { 1 })
    if ($nombre -match "(?i)WEB-?DL")      { Set-ComboValor $ui.upTipo 6 }
    elseif ($nombre -match "(?i)WEB-?Rip") { Set-ComboValor $ui.upTipo 6 }
    elseif ($nombre -match "(?i)Remux")    { Set-ComboValor $ui.upTipo $(if ($nombre -match "(?i)UHD|2160") {5} else {2}) }
    elseif ($nombre -match "(?i)BDRip|BRRip") { Set-ComboValor $ui.upTipo 7 }
    Actualizar-TipoCategoria
    # Temporada/episodio del nombre
    if ($esSerie -and $nombre -match "(?i)\bS(\d{1,3})(?:E(\d{1,4}))?") {
        $ui.upTemporada.Text = [int]$Matches[1]
        $ui.upEpisodio.Text = if ($Matches[2]) { "$([int]$Matches[2])" } else { "0" }   # sin episodio = pack
    }

    # Capturas: las _cap_*.jpg de este título. El montaje las deja en la CARPETA DE ORIGEN (no se
    # mueven aunque el MKV vaya a otra carpeta de salida).
    $baseT = [System.IO.Path]::GetFileNameWithoutExtension($nombre)
    $baseEsc = [regex]::Escape($baseT)
    # 0) Si el motor nos dio las rutas EXACTAS de las capturas (al procesar), las usamos tal cual.
    #    Es lo más fiable: funciona aunque el MKV se haya movido, tras reabrir el programa, y en
    #    packs donde la captura lleva el nombre del episodio y no el del pack.
    $caps = @(@($capturasHint) | Where-Object { $_ -and (Test-Path -LiteralPath "$_") } | Select-Object -Unique)
    # 1) Si no, buscamos por título exacto en todas las carpetas candidatas (incluida la de origen
    #    que reporta el motor y la carpeta de trabajo/salida de la interfaz).
    if ($caps.Count -eq 0) {
        $dirs = @(
            "$origenHint".Trim(),
            $(if ($video) { Split-Path $video -Parent }),
            $(Split-Path $rutaTorrent -Parent),
            "$($ui.txtCarpeta.Text)".Trim(),
            "$($ui.txtSalidaArchivo.Text)".Trim()
        ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
        $caps = @($dirs | ForEach-Object {
            Get-ChildItem -LiteralPath $_ -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match "(?i)^\.(jpg|jpeg|png)$" -and $_.BaseName -match "(?i)^$baseEsc.*_cap_\d+$" }
        } | Sort-Object FullName -Unique | ForEach-Object { $_.FullName })
    }
    # 2) Respaldo: cualquier _cap_ (solo en la carpeta del vídeo/torrent, para no mezclar títulos).
    if ($caps.Count -eq 0) {
        $dirPrinc = if ($video) { Split-Path $video -Parent } else { Split-Path $rutaTorrent -Parent }
        $caps = @(Get-ChildItem -LiteralPath $dirPrinc -File -ErrorAction SilentlyContinue |
                  Where-Object { $_.Extension -match "(?i)^\.(jpg|jpeg|png)$" -and $_.BaseName -match "(?i)_cap_\d+$" } |
                  Sort-Object Name | ForEach-Object { $_.FullName })
    }
    Construir-Capturas $caps

    # Búsqueda TMDB automática (puede fallar o no acertar; el usuario lo revisa a mano).
    if (-not [string]::IsNullOrWhiteSpace($ui.cfgTmdbKey.Text)) {
        try {
            $ta = Get-TituloAnioLimpio $baseT
            Set-Estado "Buscando «$($ta.Titulo)» en TMDB…"
            $res = Search-Tmdb $ta.Titulo $ta.Anio $script:tmdbTipo
            if ($res -and @($res).Count -gt 0) {
                Aplicar-ResultadoTmdb $res[0]
                Set-Estado "Torrent cargado. TMDB: $($res[0].Titulo) ($($res[0].Fecha)) — revisa que sea correcto."
            } else {
                Set-Estado "Torrent cargado. TMDB no encontró coincidencia automática; busca a mano."
            }
        } catch {
            Set-Estado "Torrent cargado. La búsqueda automática TMDB falló: $($_.Exception.Message)"
        }
    } else {
        Set-Estado "Torrent cargado: $($ui.upTitulo.Text). (Añade tu clave TMDB en Ajustes para autocompletar IDs.)"
    }
    if (Get-Command Construir-TabBar -ErrorAction SilentlyContinue) { Construir-TabBar }   # refrescar la etiqueta de la pestaña
}

# =========================================================================
# SUBIDAS MÚLTIPLES — cada pestaña es una subida con su propio estado completo.
# Reutilizamos los mismos controles: al cambiar de pestaña hacemos snapshot del
# estado actual y restauramos el de la pestaña destino.
# =========================================================================
$script:tabs = @()
$script:tabActual = -1
$script:logosActuales = @()

function Nuevo-EstadoSubida {
    @{
        TorrentRuta=""; Video=$null; Titulo=""; Keywords=""; Tmdb=""; Imdb=""; Tvdb=""; Mal=""
        Temporada=""; Episodio=""; Categoria=0; Tipo=0; Resolucion=0
        Mediainfo=""; Nfo=""; Descripcion=$script:descVacia
        Anon=$false; Personal=$false; AudioEd=$false; Proper=$false; DV=$false; HDR10P=$false; HDR10=$false; ModQueue=$false
        Firma=$null; Logos=@(); LogoSel=$null; Caps=@()
    }
}
function Snapshot-Subida {
    @{
        TorrentRuta="$($ui.txtTorrentSubir.Text)"; Video=$script:videoSubir
        Titulo="$($ui.upTitulo.Text)"; Keywords="$($ui.upKeywords.Text)"
        Tmdb="$($ui.upTmdb.Text)"; Imdb="$($ui.upImdb.Text)"; Tvdb="$($ui.upTvdb.Text)"; Mal="$($ui.upMal.Text)"
        Temporada="$($ui.upTemporada.Text)"; Episodio="$($ui.upEpisodio.Text)"
        Categoria=$ui.upCategoria.SelectedIndex; Tipo=$ui.upTipo.SelectedIndex; Resolucion=$ui.upResolucion.SelectedIndex
        Mediainfo="$($ui.upMediainfo.Text)"; Nfo="$($ui.upNfo.Text)"; Descripcion="$($ui.upDescripcion.Text)"
        Anon=[bool]$ui.upAnon.IsChecked; Personal=[bool]$ui.upPersonal.IsChecked; AudioEd=[bool]$ui.upAudioEd.IsChecked
        Proper=[bool]$ui.upProper.IsChecked; DV=[bool]$ui.upDV.IsChecked; HDR10P=[bool]$ui.upHDR10P.IsChecked
        HDR10=[bool]$ui.upHDR10.IsChecked; ModQueue=[bool]$ui.upModQueue.IsChecked
        Firma=$script:firmaSel; Logos=$script:logosActuales; LogoSel=$script:logoSel
        Caps=@($script:capRows | ForEach-Object { @{ Ruta=$_.Ruta; Marcada=[bool]$_.Check.IsChecked; Subida=$_.Subida } })
    }
}
function Restore-Subida($s) {
    $ui.txtTorrentSubir.Text="$($s.TorrentRuta)"; $script:videoSubir=$s.Video
    $ui.upTitulo.Text="$($s.Titulo)"; $ui.upKeywords.Text="$($s.Keywords)"
    $ui.upTmdb.Text="$($s.Tmdb)"; $ui.upImdb.Text="$($s.Imdb)"; $ui.upTvdb.Text="$($s.Tvdb)"; $ui.upMal.Text="$($s.Mal)"
    $ui.upTemporada.Text="$($s.Temporada)"; $ui.upEpisodio.Text="$($s.Episodio)"
    if ([int]$s.Categoria -ge 0)  { $ui.upCategoria.SelectedIndex=[int]$s.Categoria }
    if ([int]$s.Tipo -ge 0)       { $ui.upTipo.SelectedIndex=[int]$s.Tipo }
    if ([int]$s.Resolucion -ge 0) { $ui.upResolucion.SelectedIndex=[int]$s.Resolucion }
    $ui.upMediainfo.Text="$($s.Mediainfo)"; $ui.upNfo.Text="$($s.Nfo)"
    $ui.upDescripcion.Text=$(if ([string]::IsNullOrEmpty("$($s.Descripcion)")) { $script:descVacia } else { "$($s.Descripcion)" })
    $ui.upAnon.IsChecked=[bool]$s.Anon; $ui.upPersonal.IsChecked=[bool]$s.Personal; $ui.upAudioEd.IsChecked=[bool]$s.AudioEd
    $ui.upProper.IsChecked=[bool]$s.Proper; $ui.upDV.IsChecked=[bool]$s.DV; $ui.upHDR10P.IsChecked=[bool]$s.HDR10P
    $ui.upHDR10.IsChecked=[bool]$s.HDR10; $ui.upModQueue.IsChecked=[bool]$s.ModQueue
    Actualizar-TipoCategoria
    $script:firmaSel=$s.Firma; Refresh-FirmaSel
    Construir-Logos @($s.Logos); $script:logoSel=$s.LogoSel; Refresh-LogoSel
    Construir-Capturas @(@($s.Caps) | ForEach-Object { $_.Ruta })
    for ($i=0; $i -lt @($script:capRows).Count -and $i -lt @($s.Caps).Count; $i++) {
        $script:capRows[$i].Check.IsChecked=[bool]$s.Caps[$i].Marcada; $script:capRows[$i].Subida=$s.Caps[$i].Subida
    }
    $ui.panResultadosTmdb.Children.Clear()
    $ui.lblTorrentInfo.Visibility="Collapsed"
    # volver a modo Escribir en la descripción al cambiar de pestaña
    if ($ui.tabEscribir) { $ui.tabEscribir.IsChecked=$true }
}
function Construir-TabBar {
    if (-not $ui.panTabsSubida) { return }
    $ui.panTabsSubida.Children.Clear()
    for ($i=0; $i -lt @($script:tabs).Count; $i++) {
        $idx = $i
        # Título COMPLETO (sin truncar): las pestañas son tipo MKVToolNix, con scroll horizontal+flechas.
        $titulo = if ($i -eq $script:tabActual) { "$($ui.upTitulo.Text)" } else { "$($script:tabs[$i].Titulo)" }
        if ([string]::IsNullOrWhiteSpace($titulo)) { $titulo = "Subida $($i+1)" }

        $cont = New-Object System.Windows.Controls.Border
        $cont.CornerRadius=[System.Windows.CornerRadius]::new(7,7,0,0)   # forma de pestaña (esquinas arriba)
        $cont.Cursor=[System.Windows.Input.Cursors]::Hand
        if ($i -eq $script:tabActual) {
            # Activa: mismo color que el cuerpo + línea roja arriba; solapa el borde del cuerpo (-1 abajo) → se funde.
            $cont.Background=(Brocha "#0E0E11")
            $cont.BorderBrush=$win.Resources["AccentBrush"]; $cont.BorderThickness=[System.Windows.Thickness]::new(0,2,0,0)
            $cont.Margin=[System.Windows.Thickness]::new(0,0,3,-1)
            $cont.Padding=[System.Windows.Thickness]::new(13,7,8,8)
        } else {
            $cont.Background=(Brocha "#17171C"); $cont.BorderBrush=$win.Resources["CardBorderBrush"]
            $cont.BorderThickness=[System.Windows.Thickness]::new(1)
            $cont.Margin=[System.Windows.Thickness]::new(0,4,3,0)
            $cont.Padding=[System.Windows.Thickness]::new(13,5,8,6)
        }

        $sp = New-Object System.Windows.Controls.StackPanel; $sp.Orientation="Horizontal"
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text=$titulo; $t.VerticalAlignment="Center"; $t.FontSize=12; $t.ToolTip=$titulo
        $t.Foreground=$(if ($i -eq $script:tabActual) { [System.Windows.Media.Brushes]::White } else { $win.Resources["SubBrush"] })
        [void]$sp.Children.Add($t)
        if (@($script:tabs).Count -gt 1) {
            $x = New-Object System.Windows.Controls.TextBlock
            $x.Text=" ✕"; $x.VerticalAlignment="Center"; $x.FontSize=11; $x.Margin=[System.Windows.Thickness]::new(8,0,0,0)
            $x.Foreground=$(if ($i -eq $script:tabActual) { [System.Windows.Media.Brushes]::White } else { $win.Resources["SubBrush"] })
            $x.Cursor=[System.Windows.Input.Cursors]::Hand
            $x.Add_MouseLeftButtonUp({ param($snd,$e) $e.Handled=$true; Cerrar-Tab $idx }.GetNewClosure())
            [void]$sp.Children.Add($x)
        }
        $cont.Child=$sp
        $cont.Add_MouseLeftButtonUp({ Cambiar-Tab $idx }.GetNewClosure())
        [void]$ui.panTabsSubida.Children.Add($cont)
    }
    # Mostrar las flechas SOLO si las pestañas desbordan el ancho visible (se comprueba tras el layout).
    if ($ui.scrTabs) { $ui.scrTabs.Dispatcher.InvokeAsync({ Update-FlechasTabs }, [System.Windows.Threading.DispatcherPriority]::Loaded) | Out-Null }
}
function Update-FlechasTabs {
    if (-not $ui.scrTabs -or -not $ui.panTabsSubida) { return }
    $vp = $ui.scrTabs.ViewportWidth
    # Si el panel aún no está medido (viewport 0, p.ej. el panel está oculto), no mostramos flechas:
    # ya se recalcula cuando el panel se muestra (SizeChanged + al entrar en «Torrent y subida»).
    if ($vp -le 1) {
        $ui.btnTabIzq.Visibility = "Collapsed"; $ui.btnTabDer.Visibility = "Collapsed"; return
    }
    # ExtentWidth = ancho real del contenido desplazable (las pestañas); si supera el viewport, desbordan.
    $vis = if ($ui.scrTabs.ExtentWidth -gt ($vp + 1)) { "Visible" } else { "Collapsed" }
    $ui.btnTabIzq.Visibility = $vis
    $ui.btnTabDer.Visibility = $vis
}
function Cambiar-Tab($idx) {
    if ($idx -eq $script:tabActual -or $idx -lt 0 -or $idx -ge @($script:tabs).Count) { return }
    if ($script:tabActual -ge 0 -and $script:tabActual -lt @($script:tabs).Count) { $script:tabs[$script:tabActual] = Snapshot-Subida }
    $script:tabActual = $idx
    Restore-Subida $script:tabs[$idx]
    Construir-TabBar
}
function Nueva-Tab {
    if ($script:tabActual -ge 0 -and $script:tabActual -lt @($script:tabs).Count) { $script:tabs[$script:tabActual] = Snapshot-Subida }
    $script:tabs += (Nuevo-EstadoSubida)
    $script:tabActual = @($script:tabs).Count - 1
    Restore-Subida $script:tabs[$script:tabActual]
    Construir-TabBar
    if ($ui.scrTabs) { $ui.scrTabs.Dispatcher.InvokeAsync({ $ui.scrTabs.ScrollToRightEnd() }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null }
    Set-Estado "Nueva subida añadida. Carga o crea su torrent."
}
function Cerrar-Tab($idx) {
    if (@($script:tabs).Count -le 1) { Set-Estado "Debe quedar al menos una subida."; return }
    if ($idx -ne $script:tabActual -and $script:tabActual -ge 0) { $script:tabs[$script:tabActual] = Snapshot-Subida }
    $nueva=@(); for ($i=0;$i -lt @($script:tabs).Count;$i++){ if ($i -ne $idx){ $nueva += $script:tabs[$i] } }
    $script:tabs = $nueva
    if ($script:tabActual -gt $idx) { $script:tabActual-- }
    elseif ($script:tabActual -eq $idx -and $script:tabActual -ge @($script:tabs).Count) { $script:tabActual = @($script:tabs).Count - 1 }
    Restore-Subida $script:tabs[$script:tabActual]
    Construir-TabBar
}
function Init-Tabs {
    $script:tabs = @(Nuevo-EstadoSubida)
    $script:tabActual = 0
    Construir-TabBar
}
# Añade una subida automáticamente (al ir terminando el montaje cada vídeo): usa la pestaña actual
# si está vacía, si no crea una nueva, y carga el torrent recién creado con su vídeo y capturas.
function Anadir-SubidaAuto($torrent, $video, $origen = "", $capturas = @()) {
    $vacia = [string]::IsNullOrWhiteSpace($ui.txtTorrentSubir.Text) -and [string]::IsNullOrWhiteSpace($ui.upTitulo.Text)
    if (-not $vacia) { Nueva-Tab }
    Cargar-TorrentParaSubir $torrent $(if ($video -and (Test-Path -LiteralPath $video)) { $video } else { $null }) $origen @($capturas)
    Construir-TabBar
    Set-Estado "Subida añadida automáticamente: $(Split-Path $torrent -Leaf). Está en «Torrent y subida»." "ok"
}
$ui.btnNuevaSubida.Add_Click({ Nueva-Tab })
$ui.btnTabIzq.Add_Click({ $ui.scrTabs.ScrollToHorizontalOffset([Math]::Max(0, $ui.scrTabs.HorizontalOffset - 240)) })
$ui.btnTabDer.Add_Click({ $ui.scrTabs.ScrollToHorizontalOffset($ui.scrTabs.HorizontalOffset + 240) })
$ui.scrTabs.Add_SizeChanged({ Update-FlechasTabs })
$ui.panTabsSubida.Add_SizeChanged({ Update-FlechasTabs })
# Flechas de las pestañas de PROYECTO/AUDIO/SUBS (modo heterogéneo) — las 3 barras
$ui.btnTabProyIzq.Add_Click({ $ui.scrTabsProy.ScrollToHorizontalOffset([Math]::Max(0, $ui.scrTabsProy.HorizontalOffset - 240)) })
$ui.btnTabProyDer.Add_Click({ $ui.scrTabsProy.ScrollToHorizontalOffset($ui.scrTabsProy.HorizontalOffset + 240) })
$ui.btnTabProyAIzq.Add_Click({ $ui.scrTabsProyA.ScrollToHorizontalOffset([Math]::Max(0, $ui.scrTabsProyA.HorizontalOffset - 240)) })
$ui.btnTabProyADer.Add_Click({ $ui.scrTabsProyA.ScrollToHorizontalOffset($ui.scrTabsProyA.HorizontalOffset + 240) })
$ui.btnTabProySIzq.Add_Click({ $ui.scrTabsProyS.ScrollToHorizontalOffset([Math]::Max(0, $ui.scrTabsProyS.HorizontalOffset - 240)) })
$ui.btnTabProySDer.Add_Click({ $ui.scrTabsProyS.ScrollToHorizontalOffset($ui.scrTabsProyS.HorizontalOffset + 240) })
foreach ($scr in @($ui.scrTabsProy, $ui.scrTabsProyA, $ui.scrTabsProyS)) { $scr.Add_SizeChanged({ Update-FlechasTabsProy }) }
foreach ($pan in @($ui.panTabsProy, $ui.panTabsProyA, $ui.panTabsProyS)) { $pan.Add_SizeChanged({ Update-FlechasTabsProy }) }
# Asas de redimensión de los campos de texto (arrastrar la esquina cambia el alto).
# Acumulamos sobre el Height actual (no ActualHeight) para que el arrastre sea estable.
$ui.gripDesc.Add_DragDelta({ param($s, $e)
    $h = if ([double]::IsNaN($ui.gridDesc.Height)) { $ui.gridDesc.ActualHeight } else { $ui.gridDesc.Height }
    $n = $h + $e.VerticalChange
    if ($n -ge 120 -and $n -le 1600) { $ui.gridDesc.Height = $n }
})
$ui.gripMedia.Add_DragDelta({ param($s, $e)
    $h = if ([double]::IsNaN($ui.gridMediaInfo.Height)) { $ui.gridMediaInfo.ActualHeight } else { $ui.gridMediaInfo.Height }
    $n = $h + $e.VerticalChange
    if ($n -ge 80 -and $n -le 1600) { $ui.gridMediaInfo.Height = $n }
})
$ui.gripListado.Add_DragDelta({ param($s, $e)
    $h = if ([double]::IsNaN($ui.gridListado.Height)) { $ui.gridListado.ActualHeight } else { $ui.gridListado.Height }
    $n = $h + $e.VerticalChange
    if ($n -ge 120 -and $n -le 2000) { $ui.gridListado.Height = $n }
})

# =========================================================================
# SUBIR AL TRACKER — eventos
# =========================================================================
$ui.upCategoria.Add_SelectionChanged({ Actualizar-TipoCategoria })

$ui.btnTorrentExaminar.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Torrents (*.torrent)|*.torrent"
    if ($ui.txtCarpeta.Text -and (Test-Path -LiteralPath $ui.txtCarpeta.Text)) { $dlg.InitialDirectory = $ui.txtCarpeta.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Cargar-TorrentParaSubir $dlg.FileName }
})
$ui.btnTorrentUltimo.Add_Click({
    # Buscamos el .torrent más reciente en TODAS las carpetas relevantes: la de salida del torrent,
    # la de salida del archivo y la de trabajo (así funciona aunque cambies dónde se guarda).
    $carpetas = @("$($ui.txtSalidaTorrent.Text)", "$($ui.txtSalidaArchivo.Text)", "$($ui.txtCarpeta.Text)") |
                Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
    if ($carpetas.Count -eq 0) { Set-Estado "Elige primero una carpeta (General) o una de salida." "error"; return }
    $ult = @($carpetas | ForEach-Object { Get-ChildItem -LiteralPath $_ -File -Recurse -ErrorAction SilentlyContinue } |
             Where-Object { $_.Extension -match "(?i)^\.torrent$" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    if ($ult.Count -eq 0) { Set-Estado "No hay ningún .torrent en las carpetas conocidas." "error"; return }
    Cargar-TorrentParaSubir $ult[0].FullName
})
# --- Crear el .torrent de algo ya procesado (archivo o carpeta/pack), en segundo plano ---
# El worker dot-sourcea el módulo compartido y crea el torrent reportando progreso por $sync.
$torrentWorker = {
    param($modulo, $tipo, $ruta, $announce, $carpeta, $sync)
    try {
        . $modulo
        $cb = { param($p) $sync.Pct = $p }
        if ($tipo -eq "pack") { $sync.Resultado = New-TorrentPack $ruta $announce "" $carpeta $cb }
        else                  { $sync.Resultado = New-TorrentArchivo $ruta $announce "" $carpeta $cb }
    } catch { $sync.Error = $_.Exception.Message }
    $sync.Fin = $true
}
$script:torTimer = $null; $script:torSync = $null; $script:torPS = $null; $script:torAsync = $null; $script:torVideoHint = $null
function Start-CrearTorrentAsync($tipo, $ruta) {
    $announce = "$($ui.txtAnnounce.Text)".Trim()
    $carpeta  = "$($ui.txtSalidaTorrent.Text)".Trim()
    # Pista del vídeo para luego cargar specs/capturas (el .torrent puede acabar en otra carpeta).
    $script:torVideoHint = if ($tipo -eq "pack") {
        @(Get-ChildItem -LiteralPath $ruta -File -Recurse -ErrorAction SilentlyContinue |
          Where-Object { $_.Extension -match "(?i)^\.(mkv|mp4)$" } | Sort-Object FullName | Select-Object -First 1 -ExpandProperty FullName)
    } else { $ruta }
    $sync = [hashtable]::Synchronized(@{ Pct = 0; Resultado = $null; Error = $null; Fin = $false })
    $script:torSync = $sync
    foreach ($b in @($ui.btnCrearTorrentArchivo, $ui.btnCrearTorrentCarpeta, $ui.btnTorrentExaminar, $ui.btnTorrentUltimo)) { $b.IsEnabled = $false }
    $ui.lblTorrentInfo.Visibility = "Collapsed"
    $ui.panTorrentProgreso.Visibility = "Visible"
    $ui.barTorrentProgreso.Value = 0; $ui.lblTorrentProgresoPct.Text = "0%"
    $ui.lblTorrentProgreso.Text = "Creando torrent de $(Split-Path $ruta -Leaf)…"

    $ps = [powershell]::Create()
    [void]$ps.AddScript($torrentWorker.ToString()).AddArgument($rutaTorrentMod).AddArgument($tipo).AddArgument($ruta).AddArgument($announce).AddArgument($carpeta).AddArgument($sync)
    $script:torPS = $ps
    $script:torAsync = $ps.BeginInvoke()

    if (-not $script:torTimer) {
        $script:torTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:torTimer.Interval = [TimeSpan]::FromMilliseconds(200)
        $script:torTimer.Add_Tick({
            $s = $script:torSync
            if (-not $s) { $script:torTimer.Stop(); return }
            $ui.barTorrentProgreso.Value = [int]$s.Pct; $ui.lblTorrentProgresoPct.Text = "$([int]$s.Pct)%"
            if ($s.Fin) {
                $script:torTimer.Stop()
                try { $script:torPS.EndInvoke($script:torAsync) } catch {}
                try { $script:torPS.Dispose() } catch {}
                $script:torPS = $null; $script:torAsync = $null
                $ui.panTorrentProgreso.Visibility = "Collapsed"
                foreach ($b in @($ui.btnCrearTorrentArchivo, $ui.btnCrearTorrentCarpeta, $ui.btnTorrentExaminar, $ui.btnTorrentUltimo)) { $b.IsEnabled = $true }
                $res = $s.Resultado; $err = $s.Error; $script:torSync = $null
                if ($err) { Set-Estado "No se pudo crear el torrent: $err" "error" }
                elseif ($res) { Set-Estado "✓  Torrent creado: $(Split-Path $res -Leaf)" "ok"; Cargar-TorrentParaSubir $res $script:torVideoHint }
            }
        })
    }
    $script:torTimer.Start()
}
$ui.btnCrearTorrentArchivo.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Vídeos (*.mkv;*.mp4)|*.mkv;*.mp4|Todos|*.*"
    if ($ui.txtCarpeta.Text -and (Test-Path -LiteralPath $ui.txtCarpeta.Text)) { $dlg.InitialDirectory = $ui.txtCarpeta.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Start-CrearTorrentAsync "archivo" $dlg.FileName }
})
$ui.btnCrearTorrentCarpeta.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Carpeta del pack (con los MKV finales) para crear su .torrent"
    if ($ui.txtCarpeta.Text -and (Test-Path -LiteralPath $ui.txtCarpeta.Text)) { $dlg.SelectedPath = $ui.txtCarpeta.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Start-CrearTorrentAsync "pack" $dlg.SelectedPath }
})
$ui.btnNfoExaminar.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "NFO (*.nfo;*.txt)|*.nfo;*.txt|Todos|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $ui.upNfo.Text = $dlg.FileName }
})
$ui.btnMediaRefrescar.Add_Click({
    if ($script:videoSubir -and (Test-Path -LiteralPath $script:videoSubir)) {
        $ui.upMediainfo.Text = Get-MediaInfoTexto $script:videoSubir
    } else { Set-Estado "No hay vídeo asociado para regenerar el MediaInfo." "error" }
})
$ui.btnCapGenerar.Add_Click({ Start-CapturasAsync })
$ui.btnCapTodas.Add_Click({ foreach ($r in $script:capRows) { $r.Check.IsChecked = $true } })
$ui.btnCapNinguna.Add_Click({ foreach ($r in $script:capRows) { $r.Check.IsChecked = $false } })
$ui.btnCapAnadir.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Imágenes (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png"
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $actuales = @($script:capRows | ForEach-Object { $_.Ruta })
        Construir-Capturas (@($actuales) + @($dlg.FileNames))
    }
})
$ui.btnCapInsertar.Add_Click({
    $marcadas = @($script:capRows | Where-Object { $_.Check.IsChecked })
    if ($marcadas.Count -eq 0) { Set-Estado "Marca al menos una captura." "error"; return }
    # Solo subimos las que aún no tienen URL (cache de sesión); el resto ya está subido.
    $pend = @($marcadas | Where-Object { -not $_.Subida })
    $onDone = {
        param($urls)
        for ($i = 0; $i -lt $pend.Count -and $i -lt @($urls).Count; $i++) { $pend[$i].Subida = $urls[$i] }
        $todas = @($marcadas | ForEach-Object { $_.Subida } | Where-Object { $_ })
        Insert-ImagenesEnDescripcion $todas
        Set-Estado "$($todas.Count) captura(s) añadidas a la descripción."
    }.GetNewClosure()
    if ($pend.Count -eq 0) { & $onDone @() }
    else { Start-SubidaAsync (@($pend | ForEach-Object { $_.Ruta })) $onDone @($ui.btnCapInsertar) }
})

# --- Toolbar BBCode de la descripción (iconos como en la web) ---
# Crea un botón-icono. $deco: bold|italic|underline|strike para las letras B/I/U/S;
# $glifo=$true usa la fuente de iconos Segoe MDL2 Assets.
function New-BbBtn($contenido, $tooltip, [scriptblock]$accion, $glifo = $false, $deco = "") {
    $b = New-Object System.Windows.Controls.Button
    $b.Style = $win.Resources["IconBtn"]
    $b.ToolTip = $tooltip
    $b.Margin = [System.Windows.Thickness]::new(0, 0, 4, 4)
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $contenido
    if ($glifo) { $tb.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe MDL2 Assets"); $tb.FontSize = 15 }
    else { $tb.FontSize = 15.5; $tb.FontWeight = "SemiBold" }
    switch ($deco) {
        "bold"      { $tb.FontWeight = "Bold" }
        "italic"    { $tb.FontStyle = "Italic" }
        "underline" { $tb.TextDecorations = [System.Windows.TextDecorations]::Underline }
        "strike"    { $tb.TextDecorations = [System.Windows.TextDecorations]::Strikethrough }
    }
    $b.Content = $tb
    $b.Add_Click($accion)
    return $b
}
function Add-BbSep {
    $s = New-Object System.Windows.Controls.Border
    $s.Width = 1; $s.Height = 20; $s.Background = Brocha "#33333A"
    $s.Margin = [System.Windows.Thickness]::new(4, 0, 7, 4); $s.VerticalAlignment = "Center"
    [void]$ui.panBbToolbar.Children.Add($s)
}

# Inserta texto literal en el cursor (sin etiquetas), conservando la posición
function Insert-Texto($texto) {
    $tb = $ui.upDescripcion
    $s = $tb.SelectionStart; $l = $tb.SelectionLength
    if ($s -lt 0) { $s = $tb.Text.Length; $l = 0 }
    $tb.Text = $tb.Text.Remove($s, $l).Insert($s, $texto)
    $tb.SelectionStart = $s + $texto.Length
    $tb.SelectionLength = 0
    $tb.Focus()
}

# Crea el botón Emoji con un popup selector claro (rejilla de emojis por categorías)
function New-EmojiBtn {
    $btn = New-BbBtn ([char]0xE76E) "Emoji" { } $true

    $pop = New-Object System.Windows.Controls.Primitives.Popup
    $pop.PlacementTarget = $btn
    $pop.Placement = "Bottom"
    $pop.StaysOpen = $false
    $pop.AllowsTransparency = $true
    $pop.PopupAnimation = "Fade"

    $marco = New-Object System.Windows.Controls.Border
    $marco.Background = Brocha "#26262C"
    $marco.BorderBrush = Brocha "#3A3A42"
    $marco.BorderThickness = [System.Windows.Thickness]::new(1)
    $marco.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $marco.Padding = [System.Windows.Thickness]::new(8)
    $marco.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
    $efx = New-Object System.Windows.Media.Effects.DropShadowEffect
    $efx.BlurRadius = 14; $efx.ShadowDepth = 3; $efx.Opacity = 0.5; $efx.Color = [System.Windows.Media.Colors]::Black
    $marco.Effect = $efx

    $cats = [ordered]@{
        "Caras"     = "😀 😃 😄 😁 😆 😅 😂 🤣 😊 🙂 😉 😍 🥰 😘 😎 🤩 🤔 🤨 😐 😴 😢 😭 😡 🤯 😱 🥳 🤗 🤫 🙄 😏"
        "Gestos"    = "👍 👎 👌 ✌️ 🤞 🤟 🤙 👏 🙌 🙏 💪 👀 🫡 🤝 ✋ 👋 ☝️ 👇 👉 👈"
        "Símbolos"  = "❤️ 🧡 💛 💚 💙 💜 🖤 ⭐ 🌟 ✨ 🔥 💯 ✅ ❌ ⚠️ ❓ ❗ 💩 🎉 🎬 🍿 📺 🎵 🏆 👑 💎 ⚡ 💥 🚀 🎯"
    }

    $stack = New-Object System.Windows.Controls.StackPanel
    foreach ($cat in $cats.Keys) {
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $cat
        $lbl.Foreground = Brocha "#9A9AA6"
        $lbl.FontSize = 10; $lbl.FontWeight = "SemiBold"
        $lbl.Margin = [System.Windows.Thickness]::new(2, 4, 0, 3)
        [void]$stack.Children.Add($lbl)

        $wrap = New-Object System.Windows.Controls.WrapPanel
        $wrap.Width = 300
        foreach ($e in ($cats[$cat] -split ' ')) {
            $emoji = $e
            $eb = New-Object System.Windows.Controls.Button
            $eb.Content = $emoji
            $eb.FontSize = 17
            $eb.Width = 30; $eb.Height = 30
            $eb.Margin = [System.Windows.Thickness]::new(1)
            $eb.Background = [System.Windows.Media.Brushes]::Transparent
            $eb.BorderThickness = [System.Windows.Thickness]::new(0)
            $eb.Cursor = "Hand"
            $eb.ToolTip = $emoji
            $eb.Add_Click({ Insert-Texto $emoji; $pop.IsOpen = $false }.GetNewClosure())
            [void]$wrap.Children.Add($eb)
        }
        [void]$stack.Children.Add($wrap)
    }

    $marco.Child = $stack
    $pop.Child = $marco
    $btn.Add_Click({ $pop.IsOpen = -not $pop.IsOpen }.GetNewClosure())
    return $btn
}
$tp = $ui.panBbToolbar
[void]$tp.Children.Add((New-BbBtn "B" "Negrita"        { Insert-Bb "[b]" "[/b]" } $false "bold"))
[void]$tp.Children.Add((New-BbBtn "I" "Cursiva"        { Insert-Bb "[i]" "[/i]" } $false "italic"))
[void]$tp.Children.Add((New-BbBtn "U" "Subrayado"      { Insert-Bb "[u]" "[/u]" } $false "underline"))
[void]$tp.Children.Add((New-BbBtn "S" "Tachado"        { Insert-Bb "[s]" "[/s]" } $false "strike"))
Add-BbSep
[void]$tp.Children.Add((New-BbBtn ([char]0xEB9F) "Imagen"   { Insert-Bb "[img=350]" "[/img]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE714) "YouTube"  { Insert-Bb "[youtube]" "[/youtube]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE71B) "Enlace"   { Insert-Bb "[url=]" "[/url]" } $true))
Add-BbSep
[void]$tp.Children.Add((New-BbBtn ([char]0xE8FD) "Lista"          { Insert-Bb "[list]`n[*]" "`n[/list]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE292) "Lista numerada" { Insert-Bb "[list=1]`n[*]" "`n[/list]" } $true))
Add-BbSep
[void]$tp.Children.Add((New-BbBtn ([char]0xE790) "Color"          { Insert-Bb "[color=#ff0000]" "[/color]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE8E9) "Tamaño de letra" { Insert-Bb "[size=24]" "[/size]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE8D2) "Tipo de letra"  { Insert-Bb "[font=Arial]" "[/font]" } $true))
Add-BbSep
[void]$tp.Children.Add((New-BbBtn ([char]0xE8E4) "Alinear izquierda" { Insert-Bb "[left]" "[/left]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE8E3) "Centrar"            { Insert-Bb "[center]" "[/center]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE8E2) "Alinear derecha"    { Insert-Bb "[right]" "[/right]" } $true))
Add-BbSep
[void]$tp.Children.Add((New-BbBtn ([char]0xE134) "Cita"     { Insert-Bb "[quote]" "[/quote]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE943) "Código"   { Insert-Bb "[code]" "[/code]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xED1A) "Spoiler"  { Insert-Bb "[spoiler]" "[/spoiler]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE891) "Nota"     { Insert-Bb "[note]" "[/note]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE7BA) "Alerta"   { Insert-Bb "[alert]" "[/alert]" } $true))
[void]$tp.Children.Add((New-BbBtn ([char]0xE80A) "Tabla"    { Insert-Bb "[table]`n[tr][td]" "[/td][/tr]`n[/table]" } $true))
[void]$tp.Children.Add((New-EmojiBtn))

$ui.btnDescSubir.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Imágenes (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png"
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    Start-SubidaAsync @($dlg.FileNames) {
        param($urls)
        Insert-ImagenesEnDescripcion $urls
        Set-Estado "$(@($urls).Count) imagen(es) subidas y añadidas a la descripción."
    } @($ui.btnDescSubir)
})

# --- Vista previa de la descripción en WPF puro (FlowDocument): sin control nativo, así recorta y
#     hace scroll correctamente dentro de la tarjeta (el WebBrowser tenía bug de "airspace"). ---
function Build-PreviewDoc($bb) {
    $doc = New-Object System.Windows.Documents.FlowDocument
    $doc.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    $doc.FontSize = 14
    $doc.PagePadding = [System.Windows.Thickness]::new(12)
    $doc.Foreground = (Brocha "#E8E8EE")

    $align = "Left"; $bold = $false; $ital = $false; $under = $false; $strike = $false; $color = $null; $size = $null
    $nuevoPara = {
        $p = New-Object System.Windows.Documents.Paragraph
        $p.TextAlignment = $align
        $p.Margin = [System.Windows.Thickness]::new(0, 0, 0, 2)
        $p.LineHeight = 22
        [void]$doc.Blocks.Add($p)
        $p
    }
    $para = & $nuevoPara
    $addRun = {
        param($t)
        $r = New-Object System.Windows.Documents.Run($t)
        if ($bold) { $r.FontWeight = "Bold" }
        if ($ital) { $r.FontStyle = "Italic" }
        if ($under -or $strike) {
            $tc = New-Object System.Windows.TextDecorationCollection
            if ($under)  { foreach ($d in [System.Windows.TextDecorations]::Underline)     { [void]$tc.Add($d) } }
            if ($strike) { foreach ($d in [System.Windows.TextDecorations]::Strikethrough) { [void]$tc.Add($d) } }
            $r.TextDecorations = $tc
        }
        if ($color) { try { $r.Foreground = (Brocha $color) } catch {} }
        if ($size)  { try { $r.FontSize = [double]$size } catch {} }
        $para.Inlines.Add($r)
    }
    $addImg = {
        param($url, $w)
        try {
            $img = New-Object System.Windows.Controls.Image
            $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
            $bmp.BeginInit(); $bmp.UriSource = [Uri]$url
            if ($w -gt 0) { $bmp.DecodePixelWidth = [int]$w }
            $bmp.EndInit()                         # carga asíncrona (no bloquea)
            $img.Source = $bmp; $img.Stretch = "Uniform"
            if ($w -gt 0) { $img.Width = [double]$w } else { $img.MaxWidth = 760 }
            $img.Margin = [System.Windows.Thickness]::new(2)
            $para.Inlines.Add((New-Object System.Windows.Documents.InlineUIContainer($img)))
        } catch {}
    }

    $imgMode = $false; $imgW = 0
    $rx = [regex]'(?s)(\[/?[a-zA-Z*][^\]]*\])|([^\[]+)'
    foreach ($m in $rx.Matches("$bb")) {
        if ($m.Groups[1].Success) {
            $tl = $m.Groups[1].Value.ToLower()
            switch -regex ($tl) {
                '^\[img(=\d+)?\]$' { $imgMode = $true; $imgW = $(if ($tl -match '=(\d+)') { [int]$Matches[1] } else { 0 }) }
                '^\[/img\]$'       { $imgMode = $false }
                '^\[b\]$'          { $bold = $true }
                '^\[/b\]$'         { $bold = $false }
                '^\[i\]$'          { $ital = $true }
                '^\[/i\]$'         { $ital = $false }
                '^\[u\]$'          { $under = $true }
                '^\[/u\]$'         { $under = $false }
                '^\[s\]$'          { $strike = $true }
                '^\[/s\]$'         { $strike = $false }
                '^\[color=.+\]$'   { $color = ($tl -replace '^\[color=(.+)\]$', '$1') }
                '^\[/color\]$'     { $color = $null }
                '^\[size=\d+\]$'   { $size = ($tl -replace '\D', '') }
                '^\[/size\]$'      { $size = $null }
                '^\[center\]$'     { $align = "Center"; $para = & $nuevoPara }
                '^\[/center\]$'    { $align = "Left";   $para = & $nuevoPara }
                '^\[right\]$'      { $align = "Right";  $para = & $nuevoPara }
                '^\[/right\]$'     { $align = "Left";   $para = & $nuevoPara }
                '^\[left\]$'       { $align = "Left";   $para = & $nuevoPara }
                '^\[/left\]$'      { $align = "Left";   $para = & $nuevoPara }
                '^\[\*\]$'         { $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak)); & $addRun "• " }
                default            { }   # cualquier otra etiqueta: se ignora (se ve el texto interior)
            }
        } else {
            $txt = $m.Groups[2].Value
            if ($imgMode) {
                $u = $txt.Trim(); if ($u) { & $addImg $u $imgW }
            } else {
                $partes = $txt -replace "`r", "" -split "`n", -1
                for ($i = 0; $i -lt $partes.Count; $i++) {
                    if ($i -gt 0) { $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak)) }
                    if ($partes[$i] -ne "") { & $addRun $partes[$i] }
                }
            }
        }
    }
    return $doc
}
# Alterna el campo de descripción entre edición (TextBox) y vista previa (FlowDocument WPF),
# en el mismo sitio, como las pestañas del tracker.
function Mostrar-ModoDescripcion($preview) {
    if ($preview) {
        try { $ui.fdPreview.Document = Build-PreviewDoc $ui.upDescripcion.Text } catch {}
        $ui.panEdicionTools.Visibility = "Collapsed"
        $ui.upDescripcion.Visibility = "Collapsed"
        $ui.panPreview.Visibility = "Visible"
    } else {
        $ui.panPreview.Visibility = "Collapsed"
        $ui.upDescripcion.Visibility = "Visible"
        $ui.panEdicionTools.Visibility = "Visible"
    }
}
$ui.tabEscribir.Add_Checked({ Mostrar-ModoDescripcion $false })
$ui.tabPrevia.Add_Checked({ Mostrar-ModoDescripcion $true })

$ui.btnBuscarTmdb.Add_Click({
    $ui.panResultadosTmdb.Children.Clear()
    $ta = Get-TituloAnioLimpio $ui.upTitulo.Text
    try {
        $res = Search-Tmdb $ta.Titulo $ta.Anio $script:tmdbTipo
        if (-not $res -or @($res).Count -eq 0) {
            $t = New-Object System.Windows.Controls.TextBlock
            $t.Text = "Sin resultados para «$($ta.Titulo)»$(if ($ta.Anio) { " ($($ta.Anio))" }). Prueba a ajustar el título."
            $t.Foreground = $win.Resources["WarnBrush"]; $t.FontSize = 12; $t.TextWrapping = "Wrap"
            [void]$ui.panResultadosTmdb.Children.Add($t)
            return
        }
        foreach ($item in $res) {
            $b = New-Object System.Windows.Controls.Button
            $b.Style = $win.Resources["GhostMini"]
            $b.HorizontalAlignment = "Left"
            $b.Margin = [System.Windows.Thickness]::new(0, 0, 0, 5)
            $marcaAnim = if (@($item.Generos) -contains 16) { "  · 🎨 animación" } else { "" }
            $b.Content = "$($item.Titulo)  ·  $($item.Fecha)  ·  TMDB $($item.Id)$marcaAnim"
            $itemCap = $item
            $b.Add_Click({
                Set-Estado "Aplicando TMDB $($itemCap.Id) (IDs, keywords, logos)…"
                Aplicar-ResultadoTmdb $itemCap
                Set-Estado "Datos rellenados desde TMDB: $($itemCap.Titulo)."
            }.GetNewClosure())
            [void]$ui.panResultadosTmdb.Children.Add($b)
        }
    } catch {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = "Error en la búsqueda TMDB: $($_.Exception.Message)"
        $t.Foreground = $win.Resources["ErrBrush"]; $t.FontSize = 12; $t.TextWrapping = "Wrap"
        [void]$ui.panResultadosTmdb.Children.Add($t)
    }
})

# El estado de la subida se muestra en la barra inferior (junto al botón «Subir a HDZERO»).
function Set-SubidaEstado($texto, $tipo = "info") {
    $ui.lblEstado.Text = $texto
    $ui.lblEstado.Foreground = switch ($tipo) {
        "ok" { $win.Resources["OkBrush"] } "error" { $win.Resources["ErrBrush"] }
        "warn" { $win.Resources["WarnBrush"] } default { $win.Resources["SubBrush"] }
    }
}

# Construye la subida completa: sube capturas seleccionadas, monta el BBCode y hace el POST a la API.
function Invoke-SubidaTracker {
    # --- Validaciones ---
    $rutaTorrent = "$($ui.txtTorrentSubir.Text)"
    if ([string]::IsNullOrWhiteSpace($rutaTorrent) -or -not (Test-Path -LiteralPath $rutaTorrent)) {
        Set-SubidaEstado "Elige primero un archivo .torrent." "error"; return
    }
    $token = "$($ui.cfgTrackerToken.Text)".Trim()
    $baseUrl = "$($ui.cfgTrackerUrl.Text)".Trim().TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($token) -or [string]::IsNullOrWhiteSpace($baseUrl)) {
        Set-SubidaEstado "Falta la URL del tracker o el token de API (pestaña Ajustes / claves)." "error"; return
    }
    if ([string]::IsNullOrWhiteSpace($ui.upTitulo.Text)) { Set-SubidaEstado "El título no puede estar vacío." "error"; return }
    if ([string]::IsNullOrWhiteSpace($ui.upMediainfo.Text)) { Set-SubidaEstado "El MediaInfo está vacío." "error"; return }
    $catId = Get-ComboValor $ui.upCategoria
    $tipoId = Get-ComboValor $ui.upTipo
    $resId = Get-ComboValor $ui.upResolucion
    if (-not $catId -or -not $tipoId -or -not $resId) { Set-SubidaEstado "Elige categoría, tipo y resolución." "error"; return }
    $esTv = ($script:tmdbTipo -eq "tv")
    if ($esTv -and ([string]::IsNullOrWhiteSpace($ui.upTemporada.Text) -or [string]::IsNullOrWhiteSpace($ui.upEpisodio.Text))) {
        Set-SubidaEstado "Para series, indica número de temporada y episodio (0 = pack)." "error"; $ui.navSubida.IsChecked = $true; return
    }

    $ui.btnSubir.IsEnabled = $false
    try {
        # --- 1. Capturas marcadas que aún NO estén en la descripción: subir e insertar.
        #        (Si ya las insertaste con «Subir y añadir», no se duplican.) ---
        $marcadas = @($script:capRows | Where-Object { $_.Check.IsChecked })
        $pend = @()
        $idx = 0
        foreach ($r in $marcadas) {
            $idx++
            if (-not $r.Subida) {
                Set-SubidaEstado "Subiendo captura $idx/$($marcadas.Count)…"
                try { $r.Subida = Send-Imagen $r.Ruta } catch {
                    Set-SubidaEstado "Error subiendo $(Split-Path $r.Ruta -Leaf): $($_.Exception.Message)" "error"
                    $ui.btnSubir.IsEnabled = $true; return
                }
            }
            if ($r.Subida -and "$($ui.upDescripcion.Text)" -notlike "*$($r.Subida)*") { $pend += $r.Subida }
        }
        if ($pend.Count -gt 0) { Insert-ImagenesEnDescripcion $pend }

        # --- 2. Descripción: siempre centrada ---
        $desc = "$($ui.upDescripcion.Text)".Trim()
        if ($desc -notmatch '(?i)\[center\]') { $desc = "[center]`n$desc`n[/center]" }

        # --- 2b. Firma elegida: subir y añadir antes del último [/center] ---
        if ($script:firmaSel -and (Test-Path -LiteralPath $script:firmaSel)) {
            Set-SubidaEstado "Subiendo firma…"
            try {
                if (-not $script:firmaCacheUrl[$script:firmaSel]) { $script:firmaCacheUrl[$script:firmaSel] = Send-Imagen $script:firmaSel }
                $urlFirma = $script:firmaCacheUrl[$script:firmaSel]
                if ($urlFirma -and "$desc" -notlike "*$urlFirma*") {
                    $tagFirma = "[img=350]$urlFirma" + "[/img]"
                    $idx = $desc.LastIndexOf("[/center]")
                    if ($idx -ge 0) { $desc = $desc.Substring(0, $idx).TrimEnd() + "`n" + $tagFirma + "`n" + $desc.Substring($idx) }
                    else { $desc = $desc.TrimEnd() + "`n$tagFirma" }
                }
            } catch {
                Set-SubidaEstado "No se pudo subir la firma: $($_.Exception.Message)" "error"; $ui.btnSubir.IsEnabled = $true; return
            }
        }

        # --- 3. Construir el multipart de la API ---
        Set-SubidaEstado "Enviando a HDZERO…"
        $imdbVal = ("$($ui.upImdb.Text)" -replace "(?i)tt", "").Trim()
        $form = [ordered]@{
            name             = "$($ui.upTitulo.Text)".Trim()
            description      = $desc
            mediainfo        = "$($ui.upMediainfo.Text)"
            category_id      = "$catId"
            type_id          = "$tipoId"
            resolution_id    = "$resId"
            tmdb             = $(if ("$($ui.upTmdb.Text)".Trim()) { "$($ui.upTmdb.Text)".Trim() } else { "0" })
            imdb             = $(if ($imdbVal) { $imdbVal } else { "0" })
            tvdb             = $(if ("$($ui.upTvdb.Text)".Trim()) { "$($ui.upTvdb.Text)".Trim() } else { "0" })
            mal              = $(if ("$($ui.upMal.Text)".Trim()) { "$($ui.upMal.Text)".Trim() } else { "0" })
            keywords         = "$($ui.upKeywords.Text)".Trim()
            anonymous        = $(if ($ui.upAnon.IsChecked) { "1" } else { "0" })
            personal_release = $(if ($ui.upPersonal.IsChecked) { "1" } else { "0" })
            mod_queue_opt_in = $(if ($ui.upModQueue.IsChecked) { "1" } else { "0" })
            # Campos personalizados de HDZERO (se ignoran si la API no los lee)
            audio_editado    = $(if ($ui.upAudioEd.IsChecked) { "1" } else { "0" })
            dolby_vision     = $(if ($ui.upDV.IsChecked) { "1" } else { "0" })
            hdr10plus        = $(if ($ui.upHDR10P.IsChecked) { "1" } else { "0" })
            hdr              = $(if ($ui.upHDR10.IsChecked) { "1" } else { "0" })
            is_proper        = $(if ($ui.upProper.IsChecked) { "1" } else { "0" })
            torrent          = Get-Item -LiteralPath $rutaTorrent
        }
        if ($esTv) {
            $form.season_number = "$($ui.upTemporada.Text)".Trim()
            $form.episode_number = "$($ui.upEpisodio.Text)".Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($ui.upNfo.Text) -and (Test-Path -LiteralPath $ui.upNfo.Text)) {
            $form.nfo = Get-Item -LiteralPath $ui.upNfo.Text
        }

        $uri = "$baseUrl/api/torrents/upload?api_token=$token"
        $resp = Invoke-RestMethod -Uri $uri -Method Post -Form $form -Headers @{ "Accept" = "application/json" } -TimeoutSec 300

        $msg = "$($resp.message)$($resp.data)"
        if ("$($resp.success)" -eq "True" -or "$msg" -match "(?i)success|uploaded|éxito") {
            Set-SubidaEstado "✓  Subido correctamente. $msg" "ok"
        } else {
            Set-SubidaEstado "Respuesta del tracker: $msg" "warn"
        }
    } catch {
        $detalle = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detalle = $_.ErrorDetails.Message }
        Set-SubidaEstado "Error en la subida: $detalle" "error"
    } finally {
        $ui.btnSubir.IsEnabled = $true
    }
}

$ui.btnSubir.Add_Click({ Invoke-SubidaTracker })

# =========================================================================
# PERSISTENCIA DE AJUSTES
# =========================================================================
function ChipPersist($nombre) { $v = Get-ChipValor $nombre; if ($null -eq $v) { "ASK" } else { "$v" } }
function ComboPersist($cmb)   { $v = Get-ComboValor $cmb;   if ($null -eq $v) { "ASK" } else { "$v" } }

function Guardar-Ajustes {
    # Defensa 1: en modo test (renders headless) NUNCA tocamos los ajustes del usuario.
    if ($script:modoTest) { return }
    try {
        # Defensa 2: leemos lo ya guardado para NO pisar credenciales con campos vacíos.
        # Si un campo sensible está vacío pero antes tenía valor, conservamos el anterior
        # (un arranque que no recargó bien las claves no podrá borrarlas al cerrar).
        $prev = $null
        if (Test-Path -LiteralPath $rutaAjustes) {
            try { $prev = Get-Content -LiteralPath $rutaAjustes -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
        }
        $keep = {
            param($uiVal, $prevName)
            $u = "$uiVal"
            if (-not [string]::IsNullOrWhiteSpace($u)) { return $u }
            if ($prev -and -not [string]::IsNullOrWhiteSpace("$($prev.$prevName)")) { return "$($prev.$prevName)" }
            return $u
        }
        # Solo se PERSISTE la configuración de la pestaña «Ajustes / claves» (credenciales, host de
        # imágenes, firma) más el marcador interno de versión descartada. Los datos del PROYECTO
        # (carpeta, título, año, opciones de procesado, carpetas de salida…) NO se guardan: al
        # reabrir el programa todo arranca en blanco, sin memoria del proyecto anterior.
        $aj = [ordered]@{
            Version        = 2
            TrackerUrl     = & $keep $ui.cfgTrackerUrl.Text   "TrackerUrl"
            TrackerToken   = & $keep $ui.cfgTrackerToken.Text "TrackerToken"
            HostImg        = ChipPersist "hostImg"
            ImgbbKey       = & $keep $ui.cfgImgbbKey.Text "ImgbbKey"
            TmdbKey        = & $keep $ui.cfgTmdbKey.Text  "TmdbKey"
            Firma          = $(if ($script:firmaSel) { Split-Path $script:firmaSel -Leaf } else { "" })
            Announce       = & $keep $ui.txtAnnounce.Text "Announce"
            UltimaVersionDescartada = $script:ajusteVerDescartada
            BienvenidaVista = [bool]$script:bienvenidaVista
        }
        # Defensa 3: escritura ATÓMICA con respaldo. Escribimos a un temporal, guardamos el
        # fichero actual como .bak (red de recuperación) y luego reemplazamos. Si algo falla a
        # mitad, el .json queda intacto y siempre hay un .bak con el estado anterior.
        $tmp = "$rutaAjustes.tmp"
        $aj | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $tmp -Encoding UTF8
        if (Test-Path -LiteralPath $rutaAjustes) { Copy-Item -LiteralPath $rutaAjustes -Destination "$rutaAjustes.bak" -Force }
        Move-Item -LiteralPath $tmp -Destination $rutaAjustes -Force
    } catch {}
}

function Restaurar-Ajustes {
    if (-not (Test-Path -LiteralPath $rutaAjustes)) { return }
    try {
        $aj = Get-Content -LiteralPath $rutaAjustes -Raw -Encoding UTF8 | ConvertFrom-Json
        if ("$($aj.Version)" -ne "2") { return }
        if ($aj.UltimaVersionDescartada) { $script:ajusteVerDescartada = "$($aj.UltimaVersionDescartada)" }
        if ($null -ne $aj.BienvenidaVista) { $script:bienvenidaVista = [bool]$aj.BienvenidaVista }
        # Recuperación: si una clave sensible está vacía en el fichero principal pero el respaldo
        # .bak (versión anterior) sí la tiene, la rescatamos de ahí. Así un borrado accidental se
        # auto-repara en el siguiente arranque en vez de quedarse perdido.
        $bak = $null
        if (Test-Path -LiteralPath "$rutaAjustes.bak") {
            try { $bak = Get-Content -LiteralPath "$rutaAjustes.bak" -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
        }
        $rescatar = {
            param($name)
            $v = "$($aj.$name)"
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
            if ($bak -and "$($bak.Version)" -eq "2" -and -not [string]::IsNullOrWhiteSpace("$($bak.$name)")) { return "$($bak.$name)" }
            return $v
        }
        $chip = { param($n, $v) if ("$v" -eq "ASK") { Set-ChipValor $n $null } elseif ($null -ne $v) { Set-ChipValor $n $v } }
        # Solo se RESTAURA la configuración de «Ajustes / claves» (credenciales, host de imágenes,
        # firma). Los datos del PROYECTO no se restauran: cada apertura empieza en blanco.
        $vTrackerUrl   = & $rescatar "TrackerUrl";   if ($vTrackerUrl)   { $ui.cfgTrackerUrl.Text   = $vTrackerUrl }
        $vTrackerToken = & $rescatar "TrackerToken"; if ($vTrackerToken) { $ui.cfgTrackerToken.Text = $vTrackerToken }
        $vImgbbKey     = & $rescatar "ImgbbKey";     if ($vImgbbKey)     { $ui.cfgImgbbKey.Text     = $vImgbbKey }
        $vTmdbKey      = & $rescatar "TmdbKey";      if ($vTmdbKey)      { $ui.cfgTmdbKey.Text      = $vTmdbKey }
        if ($aj.Firma)        { $script:firmaSel = Join-Path $rutaFirmas "$($aj.Firma)" }
        & $chip "hostImg" $aj.HostImg
        $vAnnounce = & $rescatar "Announce"; if ($vAnnounce) { $ui.txtAnnounce.Text = $vAnnounce }
        $ui.swProyecto.IsChecked = $true       # los datos de proyecto van siempre (sin prompts en consola)
        $ui.swAplicarTodos.IsChecked = $true   # aplicar a todos en heterogéneo (sin preguntar por archivo)
    } catch {}
}

# =========================================================================
# LANZAMIENTO DEL SCRIPT
# =========================================================================
$ui.btnIniciar.Add_Click({
    if (-not (Test-Path -LiteralPath $rutaScript)) {
        Set-Estado "No se encuentra HDZnew.ps1 junto a esta interfaz ($rutaScript)." "error"
        return
    }
    $carpeta = $ui.txtCarpeta.Text
    if ([string]::IsNullOrWhiteSpace($carpeta) -or -not (Test-Path -LiteralPath $carpeta)) {
        Set-Estado "Selecciona una carpeta válida con los vídeos." "error"
        return
    }
    if ((Get-ChipValor "modoLote") -eq "HETEROGENEO" -and @($script:tabsProy).Count -gt 0) {
        # En modo por archivo, cada pestaña (película) necesita su título. Validamos todas.
        if ($script:tabProyActual -ge 0 -and $script:tabProyActual -lt @($script:tabsProy).Count) {
            $script:tabsProy[$script:tabProyActual].Estado = Snapshot-Proyecto
        }
        for ($i = 0; $i -lt @($script:tabsProy).Count; $i++) {
            if ([string]::IsNullOrWhiteSpace("$($script:tabsProy[$i].Estado.Titulo)")) {
                $ui.navProyecto.IsChecked = $true
                Cambiar-TabProy $i
                Set-Estado "La película «$($script:tabsProy[$i].Principal)» no tiene título. Rellénalo en su pestaña." "error"
                return
            }
        }
    } elseif ([bool]$ui.swProyecto.IsChecked -and [string]::IsNullOrWhiteSpace($ui.txtTitulo.Text)) {
        Set-Estado "Has activado los datos del proyecto: el título no puede estar vacío." "error"
        $ui.navProyecto.IsChecked = $true
        return
    }
    $filtroVal = Get-ChipValor "filtro"
    $idiomasMarcados = @($script:chipsIdiomas | Where-Object { $_.IsChecked } | ForEach-Object { "$($_.Tag)" })
    if ($filtroVal -eq "PERSONALIZADA" -and $idiomasMarcados.Count -eq 0) {
        Set-Estado "Filtro personalizado de subtítulos: marca al menos un idioma." "error"
        $ui.navSubs.IsChecked = $true
        return
    }
    # Selección de archivos del análisis
    $todosNombresScan = @($script:ultimoScan | ForEach-Object { $_.Nombre })
    $marcadosSel = @($todosNombresScan | Where-Object { $script:seleccion[$_] -ne $false })
    if ($todosNombresScan.Count -gt 0 -and $marcadosSel.Count -eq 0) {
        Set-Estado "No hay ningún archivo seleccionado para procesar." "error"
        $ui.navGeneral.IsChecked = $true
        return
    }

    $cfg = [ordered]@{}
    # Si se ha deseleccionado algún archivo, el script procesará SOLO los marcados
    if ($todosNombresScan.Count -gt 0 -and $marcadosSel.Count -lt $todosNombresScan.Count) {
        $cfg.ArchivosSeleccionados = @($marcadosSel)
    }
    $v = Get-ChipValor "modoLote";     if ($null -ne $v) { $cfg.ModoLote = $v }
    $v = Get-ChipValor "reprocesar";   if ($null -ne $v) { $cfg.ReprocesarHDZ = $v }
    $v = Get-ChipValor "capturasProy"; if ($null -ne $v) { $cfg.NumCapturas = [int]$v }
    $v = Get-ChipValor "originales";   if ($null -ne $v) { $cfg.BorrarOriginales = $v }
    $v = Get-ChipValor "sufijo";       if ($null -ne $v) { $cfg.SufijoHDZ = $v }
    $v = Get-ChipValor "torrent";      if ($null -ne $v) { $cfg.ModoTorrent = $v }
    # La URL de anuncio y el nombre del pack se envían siempre que estén rellenos.
    if (-not [string]::IsNullOrWhiteSpace($ui.txtAnnounce.Text))   { $cfg.TorrentAnnounce   = $ui.txtAnnounce.Text.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($ui.txtPackNombre.Text)) { $cfg.TorrentPackNombre = $ui.txtPackNombre.Text.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($ui.txtSalidaArchivo.Text)) { $cfg.CarpetaSalida  = $ui.txtSalidaArchivo.Text.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($ui.txtSalidaTorrent.Text)) { $cfg.CarpetaTorrent = $ui.txtSalidaTorrent.Text.Trim() }
    # Conversor de audio manual (sustituye a ModoConversionDTS). Lista global; en heterogéneo
    # se añade además ConversionesPorArchivo (por película) más abajo, que tiene prioridad.
    $convGlobal = @(Leer-ConversionesActuales | ForEach-Object {
        [ordered]@{ Index = $_.Index; Lang = $_.Lang; CodecDestino = $_.CodecDestino
                    CanalesDestino = $_.CanalesDestino; BitrateK = $_.BitrateK; MantenerOriginal = $_.MantenerOriginal }
    })
    if ($convGlobal.Count -gt 0) { $cfg.ConversionesAudioManual = $convGlobal }
    $v = Get-ChipValor "defAudio";     if ($null -ne $v) { $cfg.DefaultPreferidoAudio = $v }
    # Subs únicos: decisión POR sub (idioma+formato), una por cada sección de la GUI.
    $mapaSubs = [ordered]@{}
    foreach ($r in @($script:subsUnicosRows)) {
        $mapaSubs["$($r.Cod)|$($r.Fmt)"] = $(if ($r.RbForzado.IsChecked) { "Forzado" } else { "Completo" })
    }
    if ($mapaSubs.Count -gt 0) { $cfg.SubsUnicosPorIdioma = $mapaSubs }

    # Idiomas por pista para las pistas 'und'. Todas tienen un valor (idioma o "und"),
    # así que todas se envían y el script nunca pregunta por consola.
    $undSeleccion = @()
    foreach ($r in $script:undRows) {
        $vU = Get-ComboValor $r.Combo
        if ($null -ne $vU) {
            $undSeleccion += [ordered]@{ Archivo = $r.Archivo; Id = $r.Id; Tipo = $r.Tipo; Idioma = $vU }
        }
    }
    if ($undSeleccion.Count -gt 0) { $cfg.IdiomasUndPistas = @($undSeleccion) }
    $v = Get-ChipValor "extraerPGS";   if ($null -ne $v) { $cfg.ExtraerPGS = $v }
    $v = Get-ChipValor "conservarPGS"; if ($null -ne $v) { $cfg.DecisionConservarPGS = $v }

    if ($filtroVal -eq "PERSONALIZADA") { $cfg.IdiomasSubsMantener = @($idiomasMarcados) }
    elseif ($null -ne $filtroVal)       { $cfg.IdiomasSubsMantener = $filtroVal }

    if ([bool]$ui.swProyecto.IsChecked) {
        $esWeb = ((Get-ChipValor "origen") -ne "FISICO")
        $plat = if ($esWeb) {
            if (-not [string]::IsNullOrWhiteSpace($ui.txtPlataformaOtra.Text)) { $ui.txtPlataformaOtra.Text.Trim() }
            else { Get-CodigoPlataforma $ui.cmbPlataforma.SelectedItem }
        } else {
            if (-not [string]::IsNullOrWhiteSpace($ui.txtFormatoOtro.Text)) { $ui.txtFormatoOtro.Text.Trim() }
            else { "$($ui.cmbFormato.SelectedItem)" }
        }
        $cfg.Proyecto = [ordered]@{
            Titulo            = $ui.txtTitulo.Text.Trim()
            Ano               = $ui.txtAno.Text.Trim()
            EsSerie           = [bool]$ui.swSerie.IsChecked
            TipoOrigen        = (Get-ChipValor "origen")
            WebTipo           = (Get-ChipValor "webTipo")
            PlataformaFormato = $plat
            EtiquetasExtra    = if ($esWeb) { "" } else { $ui.txtEtiquetas.Text.Trim() }
            AplicarATodos     = [bool]$ui.swAplicarTodos.IsChecked
        }
    }

    # Identidad POR ARCHIVO (modo «Cada archivo distinto»): mapa nombreArchivo -> datos de proyecto.
    # El motor (Get-DatosProyecto) lo consulta por nombre de archivo y aplica la identidad de cada peli.
    if ((Get-ChipValor "modoLote") -eq "HETEROGENEO" -and @($script:tabsProy).Count -gt 0) {
        if ($script:tabProyActual -ge 0 -and $script:tabProyActual -lt @($script:tabsProy).Count) {
            $script:tabsProy[$script:tabProyActual].Estado = Snapshot-Proyecto
            $script:tabsProy[$script:tabProyActual].Decisiones = Snapshot-DecisionesTab
        }
        $mapaProy = [ordered]@{}
        foreach ($tb in @($script:tabsProy)) {
            $e = $tb.Estado
            $esWebE = ("$($e.Origen)" -ne "FISICO")
            $platE = if ($esWebE) {
                if (-not [string]::IsNullOrWhiteSpace($e.PlatOtra)) { "$($e.PlatOtra)".Trim() }
                elseif ([int]$e.PlatIdx -ge 0 -and [int]$e.PlatIdx -lt $ui.cmbPlataforma.Items.Count) { Get-CodigoPlataforma $ui.cmbPlataforma.Items[[int]$e.PlatIdx] }
                else { "" }
            } else {
                if (-not [string]::IsNullOrWhiteSpace($e.FmtOtro)) { "$($e.FmtOtro)".Trim() }
                elseif ([int]$e.FmtIdx -ge 0 -and [int]$e.FmtIdx -lt $ui.cmbFormato.Items.Count) { "$($ui.cmbFormato.Items[[int]$e.FmtIdx])" }
                else { "" }
            }
            $datos = [ordered]@{
                Titulo            = "$($e.Titulo)".Trim()
                Ano               = "$($e.Ano)".Trim()
                EsSerie           = [bool]$e.EsSerie
                TipoOrigen        = "$($e.Origen)"
                WebTipo           = "$($e.WebTipo)"
                PlataformaFormato = $platE
                EtiquetasExtra    = if ($esWebE) { "" } else { "$($e.Etiquetas)".Trim() }
                NumCapturas       = $(if ($null -ne $e.Capturas) { [int]$e.Capturas } else { -1 })
            }
            foreach ($arch in @($tb.Archivos)) { $mapaProy[$arch] = $datos }
        }
        if ($mapaProy.Count -gt 0) { $cfg.ProyectoPorArchivo = $mapaProy }

        # Decisiones de audio/subtítulos POR ARCHIVO (audio por defecto, filtro, PGS) + conversiones
        # de audio por película + pistas und (de TODAS las pestañas) + subs únicos. Reemplazan a las globales.
        $mapaDec = [ordered]@{}; $undTodos = @(); $mapaSU = [ordered]@{}; $mapaConv = [ordered]@{}
        foreach ($tb in @($script:tabsProy)) {
            $d = $tb.Decisiones
            if (-not $d) { continue }
            $filtroEf = if ($d.Filtro -eq "PERSONALIZADA") { @($d.FiltroLangs) } else { $d.Filtro }
            $dec = [ordered]@{
                DefAudio = $d.DefAudio; Filtro = $filtroEf
                ExtraerPGS = [bool]$d.ExtraerPGS; ConservarPGS = $d.ConservarPGS
            }
            foreach ($arch in @($tb.Archivos)) { $mapaDec[$arch] = $dec }
            # Conversiones de audio de esta película (formato motor, sin el campo GUI 'Dest').
            $convPeli = @(@($d.Conv) | Where-Object { $_ } | ForEach-Object {
                [ordered]@{ Index = $_.Index; Lang = $_.Lang; CodecDestino = $_.CodecDestino
                            CanalesDestino = $_.CanalesDestino; BitrateK = $_.BitrateK; MantenerOriginal = $_.MantenerOriginal }
            })
            if ($convPeli.Count -gt 0) { foreach ($arch in @($tb.Archivos)) { $mapaConv[$arch] = $convPeli } }
            if ($d.Und) {
                foreach ($k in $d.Und.Keys) {
                    if ($null -eq $d.Und[$k]) { continue }
                    $p = "$k" -split "\|"
                    if ($p.Count -ge 3) { $undTodos += [ordered]@{ Archivo = $p[0]; Id = $p[1]; Tipo = $p[2]; Idioma = $d.Und[$k] } }
                }
            }
            if ($d.SubsUnicos -and $d.SubsUnicos.Count -gt 0) {
                $perArch = [ordered]@{}
                foreach ($kk in $d.SubsUnicos.Keys) { $perArch["$kk"] = $d.SubsUnicos[$kk] }
                foreach ($arch in @($tb.Archivos)) { $mapaSU[$arch] = $perArch }
            }
        }
        if ($mapaDec.Count -gt 0) { $cfg.DecisionesPorArchivo = $mapaDec }
        if ($mapaConv.Count -gt 0) { $cfg.ConversionesPorArchivo = $mapaConv }       # conversiones por película
        if ($undTodos.Count -gt 0) { $cfg.IdiomasUndPistas = @($undTodos) }          # reemplaza al global
        if ($mapaSU.Count -gt 0)  { $cfg.SubsUnicosPorArchivo = $mapaSU }
    }

    try {
        $marca = Get-Date -Format "yyyyMMdd_HHmmss"
        $cfgPath = Join-Path $env:TEMP "hdz_gui_config_$marca.json"
        $cfg | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $cfgPath -Encoding UTF8

        # Fichero de progreso (lo escribe HDZnew.ps1) y bandera de visibilidad de la consola
        # (la vigila un hilo nativo dentro del propio proceso de consola). Empieza oculta = "0".
        $progPath = Join-Path $env:TEMP "hdz_gui_progreso_$marca.json"
        $flagPath = Join-Path $env:TEMP "hdz_gui_consola_$marca.flag"
        $resPath  = Join-Path $env:TEMP "hdz_gui_resultados_$marca.jsonl"
        Remove-Item -LiteralPath $progPath, $resPath -ErrorAction SilentlyContinue
        "0" | Set-Content -LiteralPath $flagPath -Encoding ASCII
        $script:rutaProgreso = $progPath
        $script:rutaFlagCons = $flagPath
        $script:rutaResultados = $resPath
        $script:resultadosLeidos = 0
        $script:consolaVisible = $false

        $esc = { param($s) "$s" -replace "'", "''" }
        $launcherPath = Join-Path $env:TEMP "hdz_gui_launcher_$marca.ps1"
        # El launcher captura su PROPIA ventana de consola (GetConsoleWindow es fiable para el
        # proceso actual) y, mediante un tipo C# (Add-Type), lanza un HILO NATIVO que la
        # muestra/oculta según el fichero-bandera. Todo el bucle vigía vive en C# para que no
        # dependa de un runspace de PowerShell (un scriptblock en un Thread crudo no lo tiene).
        # Como la consola arranca con -WindowStyle Hidden, no hay parpadeo inicial.
        $plantilla = @'
Add-Type -Namespace HdzCons -Name Win -UsingNamespace System.Threading,System.IO -MemberDefinition @"
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
public static void Start(string flagPath) {
    IntPtr hwnd = GetConsoleWindow();
    Thread t = new Thread(delegate() {
        int last = -1;
        while (true) {
            int want = 1;
            try { if (File.ReadAllText(flagPath).Trim() == "0") want = 0; } catch {}
            if (want != last) { ShowWindow(hwnd, want == 0 ? 0 : 5); last = want; }
            Thread.Sleep(200);
        }
    });
    t.IsBackground = true;
    t.Start();
}
"@
[HdzCons.Win]::Start('@@FLAG@@')
Set-Location -LiteralPath '@@CARPETA@@'
$env:HDZ_CONFIG = '@@CFG@@'
$env:HDZ_PROGRESS = '@@PROG@@'
$env:HDZ_RESULTS = '@@RES@@'
$env:HDZ_CONSOLE_FLAG = '@@FLAG@@'
& '@@SCRIPT@@'
'@
        $launcher = $plantilla `
            -replace '@@FLAG@@',    (& $esc $flagPath) `
            -replace '@@CARPETA@@', (& $esc $carpeta) `
            -replace '@@CFG@@',     (& $esc $cfgPath) `
            -replace '@@PROG@@',    (& $esc $progPath) `
            -replace '@@RES@@',     (& $esc $resPath) `
            -replace '@@SCRIPT@@',  (& $esc $rutaScript)
        Set-Content -LiteralPath $launcherPath -Value $launcher -Encoding UTF8

        $script:procMontaje = Start-Process pwsh -PassThru -WindowStyle Hidden `
            -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-File", $launcherPath)

        # UI de progreso: mostrar barra, habilitar el botón de consola, arrancar el sondeo.
        $ui.btnConsola.IsEnabled = $true
        $script:consolaVisible = $false
        Set-ContenidoConsola $false
        $ui.panProgreso.Visibility = "Visible"
        $ui.panProgreso2.Visibility = "Collapsed"
        $ui.barProgreso.Value = 0
        $ui.lblProgreso.Text = "Preparando montaje…"
        $ui.lblProgresoPct.Text = "0%"
        Iniciar-SondeoProgreso

        Guardar-Ajustes
        Set-Estado "✓  Montaje lanzado en segundo plano. Pulsa «Mostrar consola» si el script pide algo." "ok"
    } catch {
        Set-Estado "Error al lanzar: $($_.Exception.Message)" "error"
    }
})

# Sondeo del fichero de progreso que escribe HDZnew.ps1 (~3/s) + detección de fin de proceso.
function Iniciar-SondeoProgreso {
    if ($script:timerProgreso) { $script:timerProgreso.Stop() }
    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(350)
    $t.Add_Tick({
        # ¿El motor abrió la consola por su cuenta (porque necesita una respuesta)? La bandera
        # pasa a "1" sin que el usuario tocara el botón → reflejarlo en la UI para no desincronizar.
        if ($script:rutaFlagCons -and -not $script:consolaVisible -and (Test-Path -LiteralPath $script:rutaFlagCons)) {
            try {
                if ((Get-Content -LiteralPath $script:rutaFlagCons -Raw -ErrorAction Stop).Trim() -eq "1") {
                    $script:consolaVisible = $true
                    Set-ContenidoConsola $true
                    Set-Estado "⚠  El montaje pide una respuesta: la consola se ha abierto. Contéstale para continuar." "warn"
                }
            } catch {}
        }
        # Leer progreso
        if ($script:rutaProgreso -and (Test-Path -LiteralPath $script:rutaProgreso)) {
            try {
                $p = Get-Content -LiteralPath $script:rutaProgreso -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($null -ne $p) {
                    $pct = [int]$p.pct
                    if ($pct -lt 0) { $pct = 0 } elseif ($pct -gt 100) { $pct = 100 }
                    $ui.barProgreso.Value = $pct
                    $ui.lblProgresoPct.Text = "$pct%"
                    $txt = "$($p.fase)"
                    if (-not [string]::IsNullOrWhiteSpace($p.archivo)) { $txt += ": $($p.archivo)" }
                    $ui.lblProgreso.Text = $txt
                    # 2ª barra (detalle de la tarea en curso): pct2 = -1 (o ausente) → oculta.
                    $pct2 = if ($null -ne $p.pct2) { [int]$p.pct2 } else { -1 }
                    if ($pct2 -ge 0) {
                        if ($pct2 -gt 100) { $pct2 = 100 }
                        $ui.panProgreso2.Visibility = "Visible"
                        $ui.barProgreso2.Value = $pct2
                        $ui.lblProgreso2Pct.Text = "$pct2%"
                        if (-not [string]::IsNullOrWhiteSpace($p.fase2)) { $ui.lblProgreso2.Text = "$($p.fase2)" }
                    } else {
                        $ui.panProgreso2.Visibility = "Collapsed"
                    }
                }
            } catch {}
        }
        # Resultados nuevos (torrents creados) → crear una pestaña de subida por cada uno.
        # Procesamos UNO por tick para no bloquear la UI (cada carga hace MediaInfo + TMDB).
        if ($script:rutaResultados -and (Test-Path -LiteralPath $script:rutaResultados)) {
            try {
                $lineas = @(Get-Content -LiteralPath $script:rutaResultados -Encoding UTF8 -ErrorAction Stop)
                if ($lineas.Count -gt $script:resultadosLeidos) {
                    $linea = $lineas[$script:resultadosLeidos]
                    $script:resultadosLeidos++
                    $r = $linea | ConvertFrom-Json
                    if ($r -and $r.torrent -and (Test-Path -LiteralPath $r.torrent)) {
                        Anadir-SubidaAuto $r.torrent "$($r.video)" "$($r.origen)" @($r.capturas)
                    }
                }
            } catch {}
        }
        # ¿Terminó el proceso de consola?
        if ($script:procMontaje -and $script:procMontaje.HasExited) {
            $script:timerProgreso.Stop()
            $ui.btnConsola.IsEnabled = $false
            $script:consolaVisible = $false
            Set-ContenidoConsola $false
            $codigo = $script:procMontaje.ExitCode
            $ui.barProgreso.Value = 100
            $ui.lblProgresoPct.Text = "100%"
            $ui.lblProgreso.Text = "Proceso finalizado."
            $ui.panProgreso2.Visibility = "Collapsed"
            Set-Estado "✓  El montaje ha finalizado. Revisa la carpeta y el log." "ok"
            $script:procMontaje = $null
        }
    })
    $script:timerProgreso = $t
    $t.Start()
}

# Contenido del botón de consola con iconos de "campo de contraseña": ojo (E7B3) para mostrar,
# ojo tachado (ED1A) para ocultar. Fuente Segoe Fluent Icons (Win11) para el glifo tachado nítido.
function Set-ContenidoConsola($consolaVisible) {
    if (-not $ui.btnConsola) { return }
    $glifo = if ($consolaVisible) { [char]0xED1A } else { [char]0xE7B3 }
    $rotulo = if ($consolaVisible) { "  Ocultar consola" } else { "  Mostrar consola" }
    $tb = New-Object System.Windows.Controls.TextBlock
    $rg = New-Object System.Windows.Documents.Run ([string]$glifo)
    $rg.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe Fluent Icons, Segoe MDL2 Assets")
    $rt = New-Object System.Windows.Documents.Run $rotulo
    [void]$tb.Inlines.Add($rg); [void]$tb.Inlines.Add($rt)
    $ui.btnConsola.Content = $tb
}
# Mostrar/ocultar la ventana de consola del montaje en curso (vía fichero-bandera).
$ui.btnConsola.Add_Click({
    if (-not $script:rutaFlagCons) { return }
    $script:consolaVisible = -not $script:consolaVisible
    $txt = if ($script:consolaVisible) { "1" } else { "0" }
    try { Set-Content -LiteralPath $script:rutaFlagCons -Value $txt -Encoding ASCII } catch {}
    Set-ContenidoConsola $script:consolaVisible
})

# Dejar el programa en blanco: borra carpeta, análisis y todas las opciones marcadas,
# como recién abierto. No toca un montaje en curso ni las credenciales de Ajustes.
function Reset-Todo {
    # General / carpeta / análisis
    $ui.txtCarpeta.Text = ""
    $script:ultimoScan = @()
    $script:seleccion = @{}
    $script:carpetaScan = $null
    $ui.panListado.Children.Clear()
    $ui.lblResumen.Text = ""
    Reset-Adaptacion

    # Todos los grupos de chips de procesamiento a su opción por defecto (índice 0).
    # Se excluye "hostImg" porque es configuración (pestaña Ajustes), no una selección.
    foreach ($nombre in @($script:gruposChips.Keys)) {
        if ($nombre -eq "hostImg") { continue }
        $lista = $script:gruposChips[$nombre]
        if (@($lista).Count -gt 0) { $lista[0].Radio.IsChecked = $true }
    }

    # Datos del proyecto (siempre activos: la GUI los define y los envía, sin prompts en consola)
    $ui.swProyecto.IsChecked = $true
    $ui.swSerie.IsChecked = $false
    $ui.swAplicarTodos.IsChecked = $true
    $ui.txtTitulo.Text = ""; $ui.txtAno.Text = ""
    $script:tituloAutoVal = ""; $script:anoAutoVal = ""; $script:serieAutoVal = $null; $script:carpetaIdentificada = ""
    $ui.txtPlataformaOtra.Text = ""; $ui.txtFormatoOtro.Text = ""; $ui.txtEtiquetas.Text = ""
    if ($ui.cmbPlataforma.Items.Count -gt 0) { $ui.cmbPlataforma.SelectedIndex = 0 }
    if ($ui.cmbFormato.Items.Count -gt 0) { $ui.cmbFormato.SelectedIndex = 0 }

    # Torrent (la URL de anuncio NO se borra: es configuración persistente de Ajustes)
    $ui.txtPackNombre.Text = ""
    $ui.txtSalidaArchivo.Text = ""; $ui.txtSalidaTorrent.Text = ""

    # Subida al tracker
    $ui.txtTorrentSubir.Text = ""
    $ui.lblTorrentInfo.Text = ""
    $ui.upTitulo.Text = ""; $ui.upKeywords.Text = ""
    $ui.upTmdb.Text = ""; $ui.upImdb.Text = ""; $ui.upTvdb.Text = ""; $ui.upMal.Text = ""
    $ui.upTemporada.Text = ""; $ui.upEpisodio.Text = ""
    $ui.upMediainfo.Text = ""; $ui.upNfo.Text = ""
    if ($ui.upCategoria.Items.Count -gt 0) { $ui.upCategoria.SelectedIndex = 0 }
    if ($ui.upTipo.Items.Count -gt 0) { $ui.upTipo.SelectedIndex = 0 }
    if ($ui.upResolucion.Items.Count -gt 0) { $ui.upResolucion.SelectedIndex = 0 }
    foreach ($cb in @("upAnon","upPersonal","upAudioEd","upProper","upDV","upHDR10P","upHDR10","upModQueue")) {
        if ($ui[$cb]) { $ui[$cb].IsChecked = $false }
    }
    $ui.panResultadosTmdb.Children.Clear()
    $ui.panLogos.Children.Clear()
    $ui.cardLogos.Visibility = "Collapsed"
    $ui.panCapturas.Children.Clear()
    $script:firmaSel = $null
    Refresh-FirmaSel
    Reset-Descripcion
    if (Get-Command Init-Tabs -ErrorAction SilentlyContinue) { Init-Tabs }   # una sola subida en blanco

    # Refrescos de visibilidad dependientes
    Actualizar-Proyecto
    Actualizar-Filtro
    Actualizar-Torrent
    Actualizar-TipoCategoria
    $ui.panImgbbKey.Visibility = if ((Get-ChipValor "hostImg") -eq "IMGBB") { "Visible" } else { "Collapsed" }

    # Si no hay montaje en curso, ocultar la barra de progreso
    if (-not ($script:procMontaje -and -not $script:procMontaje.HasExited)) {
        $ui.panProgreso.Visibility = "Collapsed"
        $ui.btnConsola.IsEnabled = $false
    }

    Set-Estado "Programa reiniciado. Elige una carpeta de vídeos para empezar."
}

# =========================================================================
# DIÁLOGO MODAL CON EL ESTILO DE LA GUI (sustituye a los MessageBox nativos)
# =========================================================================
# Devuelve $true si el usuario confirma (Sí / botón principal), $false si cancela.
# Si -BotonNo queda vacío, se muestra un único botón (modo "aceptar").
function Show-DialogoHDZ {
    param(
        [string]$Titulo,
        [string]$Mensaje,
        [string]$Icono   = "",
        [string]$BotonSi = "Sí",
        [string]$BotonNo = "No",
        [System.Windows.Window]$Owner = $null
    )
    $soloOk = [string]::IsNullOrEmpty($BotonNo)
    $xamlDlg = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        SizeToContent="Height" Width="470" ResizeMode="NoResize" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display" TextOptions.TextRenderingMode="ClearType">
  <Window.Resources>
    <Style x:Key="P" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="13.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#D92B2B" CornerRadius="9" Padding="22,10">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#EF4444"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#A81E1E"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="G" TargetType="Button">
      <Setter Property="Foreground" Value="#C8C8CE"/>
      <Setter Property="FontSize" Value="13.5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#1A1A1F" BorderBrush="#2E2E35" BorderThickness="1" CornerRadius="9" Padding="20,10">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#D92B2B"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Border Background="#131316" BorderBrush="#2E2E35" BorderThickness="1" CornerRadius="14" Margin="16">
    <Border.Effect><DropShadowEffect BlurRadius="30" ShadowDepth="7" Opacity="0.55" Color="#000000"/></Border.Effect>
    <StackPanel Margin="28,24,28,22">
      <DockPanel x:Name="barra" LastChildFill="True" Margin="0,0,0,16">
        <TextBlock x:Name="lblIcono" FontSize="22" VerticalAlignment="Center" Margin="0,0,12,0" DockPanel.Dock="Left"/>
        <TextBlock x:Name="lblTitulo" Foreground="#EDEDEF" FontSize="16.5" FontWeight="SemiBold" VerticalAlignment="Center" TextWrapping="Wrap"/>
      </DockPanel>
      <ScrollViewer MaxHeight="360" VerticalScrollBarVisibility="Auto" Margin="0,0,0,24">
        <TextBlock x:Name="lblCuerpo" Foreground="#9C9CA8" FontSize="13.5" TextWrapping="Wrap" LineHeight="21" Margin="0,0,8,0"/>
      </ScrollViewer>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="btnNo" Style="{StaticResource G}" Margin="0,0,10,0"/>
        <Button x:Name="btnSi" Style="{StaticResource P}"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>
'@
    $dlg = [Windows.Markup.XamlReader]::Parse($xamlDlg)
    $g = @{}
    foreach ($n in @("barra","lblIcono","lblTitulo","lblCuerpo","btnSi","btnNo")) { $g[$n] = $dlg.FindName($n) }
    $g.lblIcono.Text  = $Icono
    if (-not $Icono) { $g.lblIcono.Visibility = "Collapsed" }
    $g.lblTitulo.Text = $Titulo
    $g.lblCuerpo.Text = $Mensaje
    $g.btnSi.Content  = $BotonSi
    $g.btnSi.IsDefault = $true
    if ($soloOk) {
        $g.btnNo.Visibility = "Collapsed"
    } else {
        $g.btnNo.Content = $BotonNo
        $g.btnNo.IsCancel = $true
        $g.btnNo.Add_Click({ $dlg.DialogResult = $false }.GetNewClosure())
    }
    $g.btnSi.Add_Click({ $dlg.DialogResult = $true }.GetNewClosure())
    # Permitir arrastrar la ventana sin barra de título (clic sobre la cabecera).
    $g.barra.Add_MouseLeftButtonDown({ try { $dlg.DragMove() } catch {} }.GetNewClosure())
    if ($Owner) {
        $dlg.Owner = $Owner
        $dlg.WindowStartupLocation = "CenterOwner"
    } else {
        $dlg.WindowStartupLocation = "CenterScreen"
    }
    $res = $dlg.ShowDialog()
    return [bool]$res
}

# =========================================================================
# Ventana de BIENVENIDA (primer arranque): pide las credenciales del usuario.
# Las claves se guardan SOLO en %APPDATA% (nunca viajan con el programa). Se
# muestra una vez, cuando aún no hay token de tracker configurado.
# =========================================================================
function Show-Bienvenida {
    param([System.Windows.Window]$Owner = $null)
    $xamlB = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        SizeToContent="Height" Width="580" ResizeMode="NoResize" ShowInTaskbar="False"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display" TextOptions.TextRenderingMode="ClearType">
  <Window.Resources>
    <Style x:Key="In" TargetType="TextBox">
      <Setter Property="Background" Value="#0E0E11"/>
      <Setter Property="Foreground" Value="#EDEDEF"/>
      <Setter Property="CaretBrush" Value="#EDEDEF"/>
      <Setter Property="BorderBrush" Value="#242429"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#D92B2B"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="P" TargetType="Button">
      <Setter Property="Foreground" Value="White"/><Setter Property="FontSize" Value="13.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" Background="#D92B2B" CornerRadius="9" Padding="22,10"><ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/></Border>
          <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#EF4444"/></Trigger></ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>
    <Style x:Key="G" TargetType="Button">
      <Setter Property="Foreground" Value="#C8C8CE"/><Setter Property="FontSize" Value="13.5"/>
      <Setter Property="Cursor" Value="Hand"/><Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" Background="#1A1A1F" BorderBrush="#2E2E35" BorderThickness="1" CornerRadius="9" Padding="20,10"><ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/></Border>
          <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#D92B2B"/><Setter Property="Foreground" Value="White"/></Trigger></ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>
    <Style x:Key="Lb" TargetType="TextBlock"><Setter Property="Foreground" Value="#EDEDEF"/><Setter Property="FontSize" Value="12.5"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Margin" Value="2,0,0,5"/></Style>
    <Style x:Key="Hn" TargetType="TextBlock"><Setter Property="Foreground" Value="#9C9CA8"/><Setter Property="FontSize" Value="11.5"/><Setter Property="Margin" Value="2,4,0,14"/><Setter Property="TextWrapping" Value="Wrap"/></Style>
  </Window.Resources>
  <Border Background="#131316" BorderBrush="#2E2E35" BorderThickness="1" CornerRadius="14" Margin="16">
    <Border.Effect><DropShadowEffect BlurRadius="30" ShadowDepth="7" Opacity="0.55" Color="#000000"/></Border.Effect>
    <StackPanel Margin="30,26,30,24">
      <StackPanel x:Name="barra" Margin="0,0,0,4">
        <TextBlock Text="Bienvenido a HD ZERO Studio" Foreground="#EDEDEF" FontSize="18" FontWeight="SemiBold"/>
        <TextBlock Text="Configura tus credenciales para empezar." Foreground="#9C9CA8" FontSize="13" Margin="0,2,0,0"/>
      </StackPanel>
      <Border Background="#101013" BorderBrush="#242429" BorderThickness="1" CornerRadius="9" Padding="12,9" Margin="0,14,0,18">
        <TextBlock Foreground="#9C9CA8" FontSize="12" TextWrapping="Wrap"
          Text="🔒  Se guardan SOLO en tu equipo (%APPDATA%). Nunca se comparten ni viajan con el programa. Podrás cambiarlas cuando quieras en «Ajustes / claves»."/>
      </Border>

      <TextBlock Style="{StaticResource Lb}" Text="URL del tracker"/>
      <TextBox x:Name="bUrl" Style="{StaticResource In}"/>
      <TextBlock Style="{StaticResource Hn}" Text="Por defecto: https://hdzero.org"/>

      <TextBlock Style="{StaticResource Lb}" Text="Token de la API (api_token)"/>
      <TextBox x:Name="bTok" Style="{StaticResource In}"/>
      <TextBlock Style="{StaticResource Hn}" Text="En HDZERO → tu perfil → Configuración → Seguridad."/>

      <TextBlock Style="{StaticResource Lb}" Text="URL de anuncio (announce, con tu passkey)"/>
      <TextBox x:Name="bAnn" Style="{StaticResource In}"/>
      <TextBlock Style="{StaticResource Hn}" Text="La que usas para subir torrents (incluye tu passkey personal)."/>

      <TextBlock Style="{StaticResource Lb}" Text="Clave o token de TMDB (opcional)"/>
      <TextBox x:Name="bTmdb" Style="{StaticResource In}"/>
      <TextBlock Style="{StaticResource Hn}" Text="Para la búsqueda automática. themoviedb.org → Ajustes → API. Puedes dejarlo vacío."/>

      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,6,0,0">
        <Button x:Name="bSkip" Style="{StaticResource G}" Content="Ahora no" Margin="0,0,10,0"/>
        <Button x:Name="bSave" Style="{StaticResource P}" Content="Guardar y empezar"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>
'@
    $b = [Windows.Markup.XamlReader]::Parse($xamlB)
    $g = @{}
    foreach ($n in @("barra","bUrl","bTok","bAnn","bTmdb","bSkip","bSave")) { $g[$n] = $b.FindName($n) }
    # Prefijar con lo que ya hubiera (normalmente vacío en un PC nuevo).
    $g.bUrl.Text  = if (-not [string]::IsNullOrWhiteSpace($ui.cfgTrackerUrl.Text)) { $ui.cfgTrackerUrl.Text } else { "https://hdzero.org" }
    $g.bTok.Text  = "$($ui.cfgTrackerToken.Text)"
    $g.bAnn.Text  = "$($ui.txtAnnounce.Text)"
    $g.bTmdb.Text = "$($ui.cfgTmdbKey.Text)"
    $g.barra.Add_MouseLeftButtonDown({ try { $b.DragMove() } catch {} }.GetNewClosure())
    $g.bSave.IsDefault = $true
    $g.bSave.Add_Click({
        $ui.cfgTrackerUrl.Text   = "$($g.bUrl.Text)".Trim()
        $ui.cfgTrackerToken.Text = "$($g.bTok.Text)".Trim()
        $ui.txtAnnounce.Text     = "$($g.bAnn.Text)".Trim()
        $ui.cfgTmdbKey.Text      = "$($g.bTmdb.Text)".Trim()
        $script:bienvenidaVista = $true
        Guardar-Ajustes
        $b.DialogResult = $true
    }.GetNewClosure())
    $g.bSkip.IsCancel = $true
    $g.bSkip.Add_Click({
        $script:bienvenidaVista = $true   # no volver a molestar; se configura en Ajustes
        Guardar-Ajustes
        $b.DialogResult = $false
    }.GetNewClosure())
    if ($Owner) { $b.Owner = $Owner; $b.WindowStartupLocation = "CenterOwner" }
    [void]$b.ShowDialog()
}

$ui.btnReset.Add_Click({
    $ok = Show-DialogoHDZ -Owner $win -Icono "🗑" -Titulo "Empezar de Zero" `
        -Mensaje "Se borrará la carpeta seleccionada y todas las opciones marcadas, dejando el programa como recién abierto.`n`nNo se borra ningún archivo del disco ni tus credenciales de Ajustes.`n`n¿Continuar?" `
        -BotonSi "Sí, empezar de Zero" -BotonNo "Cancelar"
    if ($ok) { Reset-Todo }
})

# =========================================================================
# ARRANQUE
# =========================================================================
Add-Type -Namespace HdzNative -Name Dwm -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("dwmapi.dll")]
public static extern int DwmSetWindowAttribute(System.IntPtr hwnd, int attr, ref int value, int size);
'@
# API para fijar el icono NATIVO de la ventana (WM_SETICON) con los frames reales del .ico.
# Es más fiable que Window.Icon (que escalaba un frame pequeño → icono diminuto en la barra).
Add-Type -Namespace HdzNative -Name IconApi -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Auto)]
public static extern System.IntPtr LoadImage(System.IntPtr hinst, string lpszName, uint uType, int cx, int cy, uint fuLoad);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern System.IntPtr SendMessage(System.IntPtr hWnd, int Msg, System.IntPtr wParam, System.IntPtr lParam);
'@
# Estado del proceso de montaje en curso y de su consola (oculta por defecto).
$script:procMontaje   = $null
$script:rutaFlagCons  = $null   # fichero-bandera: "1" visible / "0" oculta
$script:rutaProgreso  = $null   # JSON de progreso que escribe HDZnew.ps1
$script:consolaVisible = $false
$script:timerProgreso = $null

$win.Add_SourceInitialized({
    try {
        $h = (New-Object System.Windows.Interop.WindowInteropHelper($win)).Handle
        $oscuro = 1
        [void][HdzNative.Dwm]::DwmSetWindowAttribute($h, 20, [ref]$oscuro, 4)
        [void][HdzNative.Dwm]::DwmSetWindowAttribute($h, 19, [ref]$oscuro, 4)
        # Icono nativo desde HDZ.ico: frame de 64 px para la barra (se reescala nítido) y de 16
        # para la barra de título. Así el icono ya no sale diminuto.
        $rutaIco = Join-Path $PSScriptRoot "HDZ.ico"
        if (Test-Path -LiteralPath $rutaIco) {
            $IMAGE_ICON = 1; $LR_LOADFROMFILE = 0x10; $WM_SETICON = 0x80
            $hBig   = [HdzNative.IconApi]::LoadImage([IntPtr]::Zero, $rutaIco, $IMAGE_ICON, 64, 64, $LR_LOADFROMFILE)
            $hSmall = [HdzNative.IconApi]::LoadImage([IntPtr]::Zero, $rutaIco, $IMAGE_ICON, 16, 16, $LR_LOADFROMFILE)
            if ($hBig   -ne [IntPtr]::Zero) { [void][HdzNative.IconApi]::SendMessage($h, $WM_SETICON, [IntPtr]1, $hBig) }
            if ($hSmall -ne [IntPtr]::Zero) { [void][HdzNative.IconApi]::SendMessage($h, $WM_SETICON, [IntPtr]0, $hSmall) }
        }
    } catch {}
})
# Al cargar: como el proceso arranca oculto (vía .vbs), el botón de la barra de tareas a veces no
# se registra hasta que la ventana cambia de estado. Forzamos el registro y traemos la ventana al
# frente recreando el botón (ShowInTaskbar off→on) y activándola.
# --- Comprobación de actualizaciones (GitHub) — no bloqueante ---
function Get-VersionLocal {
    $vf = Join-Path $PSScriptRoot "VERSION.txt"
    if (Test-Path -LiteralPath $vf) { try { return (Get-Content -LiteralPath $vf -Raw).Trim() } catch {} }
    return $script:HDZVersion
}
if ($ui.lblVersion) { $ui.lblVersion.Text = "v$(Get-VersionLocal)" }
$script:updTimer = $null; $script:updSync = $null; $script:updPS = $null; $script:updAsync = $null
function Iniciar-ChequeoActualizacion {
    if ($script:modoTest) { return }
    $url = $null
    try {
        $cfgPath = Join-Path $PSScriptRoot "HDZ-update.config.json"
        if (-not (Test-Path -LiteralPath $cfgPath)) { return }
        $cfg = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $repo = "$($cfg.repo)"; $rama = "$($cfg.rama)"; if (-not $rama) { $rama = "main" }
        if ([string]::IsNullOrWhiteSpace($repo) -or $repo -like "*USUARIO/REPO*") { return }   # repo sin configurar → no se busca
        $url = "https://raw.githubusercontent.com/$repo/$rama/version.json"
    } catch { return }

    $sync = [hashtable]::Synchronized(@{ Hecho = $false; Json = $null })
    $script:updSync = $sync
    $worker = {
        param($url, $sync)
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $sync.Json = Invoke-RestMethod -Uri $url -UseBasicParsing -TimeoutSec 20
        } catch {}
        $sync.Hecho = $true
    }
    $ps = [powershell]::Create()
    [void]$ps.AddScript($worker.ToString()).AddArgument($url).AddArgument($sync)
    $script:updPS = $ps; $script:updAsync = $ps.BeginInvoke()

    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(500)
    $t.Add_Tick({
        $s = $script:updSync
        if (-not $s -or -not $s.Hecho) { return }
        $script:updTimer.Stop()
        try { $script:updPS.EndInvoke($script:updAsync) } catch {}
        try { $script:updPS.Dispose() } catch {}
        $j = $s.Json; $script:updSync = $null
        if (-not $j -or -not $j.version) { return }
        $verRemota = "$($j.version)"; $verLocal = Get-VersionLocal
        $nueva = $false
        try { $nueva = ([version]$verRemota) -gt ([version]$verLocal) } catch { $nueva = ($verRemota -ne $verLocal) }
        if (-not $nueva) { return }
        if ("$($script:ajusteVerDescartada)" -eq $verRemota) { return }   # ya dijo "no" a esta versión
        $msg = "Hay una versión nueva de HDZ Studio: $verRemota (tienes $verLocal)."
        if ("$($j.notas)") { $msg += "`n`nNovedades:`n$($j.notas)" }
        $msg += "`n`n¿Actualizar ahora? El programa se cerrará y se reabrirá ya actualizado."
        $r = Show-DialogoHDZ -Owner $win -Icono "⬇️" -Titulo "Actualización disponible" -Mensaje $msg `
            -BotonSi "Actualizar ahora" -BotonNo "Ahora no"
        if ($r) {
            $upd = Join-Path $PSScriptRoot "instalacion\Actualizar-HDZStudio.ps1"
            if (Test-Path -LiteralPath $upd) {
                # El actualizador sobrescribe los .ps1 en la carpeta del programa. Si esa carpeta NO
                # es escribible sin permisos (p.ej. instalado en "Archivos de programa"), lo lanzamos
                # ELEVADO (-Verb RunAs) para que pueda escribir; si es una carpeta de usuario (ZIP en
                # Documentos/Escritorio) va normal, sin molestar con UAC.
                $necesitaAdmin = $false
                try {
                    $pruebaW = Join-Path $PSScriptRoot ".hdz_wtest_$PID.tmp"
                    [System.IO.File]::WriteAllText($pruebaW, "x")
                    Remove-Item -LiteralPath $pruebaW -Force -ErrorAction SilentlyContinue
                } catch { $necesitaAdmin = $true }
                $argsUpd = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$upd`"")
                try {
                    if ($necesitaAdmin) {
                        Start-Process "powershell.exe" -Verb RunAs -ArgumentList $argsUpd
                    } else {
                        Start-Process "powershell.exe" -ArgumentList $argsUpd
                    }
                    $win.Close()
                } catch {
                    Set-Estado "Actualización cancelada: no se concedieron permisos de administrador." "warn"
                }
            } else { Set-Estado "No encuentro el actualizador (instalacion\Actualizar-HDZStudio.ps1)." "error" }
        } else {
            $script:ajusteVerDescartada = $verRemota   # no volver a avisar de ESTA versión
            Guardar-Ajustes
        }
    })
    $script:updTimer = $t; $t.Start()
}

$win.Add_Loaded({
    try {
        $win.ShowInTaskbar = $false
        $win.ShowInTaskbar = $true
        if ($win.WindowState -eq [System.Windows.WindowState]::Minimized) { $win.WindowState = [System.Windows.WindowState]::Normal }
        [void]$win.Activate()
        $win.Topmost = $true; $win.Topmost = $false
    } catch {}
    # Primer arranque: si aún no hay token de tracker y no se mostró antes, guiamos al usuario con
    # la ventana de bienvenida para que ponga SUS credenciales (se guardan solo en su %APPDATA%).
    if (-not $script:modoTest -and -not $script:bienvenidaVista -and [string]::IsNullOrWhiteSpace($ui.cfgTrackerToken.Text)) {
        try { Show-Bienvenida -Owner $win } catch {}
    }
    # Buscar actualizaciones en segundo plano (no molesta si no hay repo configurado o no hay novedad).
    try { $win.Dispatcher.InvokeAsync([action]{ Iniciar-ChequeoActualizacion }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null } catch {}
})
$win.Add_Closing({ Guardar-Ajustes })

# Logo del tracker (junto a la GUI); si no existe, texto de respaldo «HDZERO».
# Acepta el logo nuevo (incluye «STUDIO») o el antiguo como reserva.
$rutaLogo = @(
    (Join-Path $PSScriptRoot "HDZero New logo.png"),
    (Join-Path $PSScriptRoot "hdz_logo_r.png")
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if ($rutaLogo) {
    try {
        $bmpLogo = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmpLogo.BeginInit()
        $bmpLogo.UriSource = New-Object System.Uri($rutaLogo)
        # SIN DecodePixelHeight y SIN BitmapScalingMode a propósito: la imagen se mantiene a
        # resolución nativa y se muestra lo más pura posible (escalado por defecto de WPF, sin
        # reducciones encadenadas ni suavizado extra que restaban definición).
        $bmpLogo.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmpLogo.EndInit()
        $ui.imgLogo.Source = $bmpLogo
        $ui.imgLogo.Visibility = "Visible"
        $ui.panLogoTexto.Visibility = "Collapsed"
        $script:logoCargado = $true
    } catch {}
}

# Icono de la ventana (y de la barra de tareas) = HDZ.ico cuadrado de marca.
# OJO: NO usar BitmapImage(UriSource=.ico): coge el frame de 16 px (¡el más pequeño!) y la barra
# de tareas lo escala → icono diminuto. Hay que abrir el decodificador y usar el frame MAYOR (256).
$rutaIco = Join-Path $PSScriptRoot "HDZ.ico"
if (Test-Path -LiteralPath $rutaIco) {
    try {
        $decIco = [System.Windows.Media.Imaging.BitmapDecoder]::Create([Uri]$rutaIco,
                  [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
                  [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $win.Icon = ($decIco.Frames | Sort-Object PixelWidth -Descending)[0]
    } catch {}
}

Restaurar-Ajustes
if ([string]::IsNullOrWhiteSpace($ui.txtCarpeta.Text)) { $ui.txtCarpeta.Text = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ui.cfgTrackerUrl.Text)) { $ui.cfgTrackerUrl.Text = "https://hdzero.org" }
Actualizar-Proyecto
Actualizar-Filtro
Actualizar-Torrent
Actualizar-ModoLote
Actualizar-TipoCategoria
Construir-Firmas
if ([string]::IsNullOrWhiteSpace($ui.upDescripcion.Text)) { Reset-Descripcion }
Init-Tabs
$ui.panImgbbKey.Visibility = if ((Get-ChipValor "hostImg") -eq "IMGBB") { "Visible" } else { "Collapsed" }
Set-Estado "Elige la carpeta de los vídeos: las opciones se adaptarán a lo que se detecte."

if ($script:modoTest) {
    # Modo test: carpeta opcional por variable de entorno, escaneo síncrono y
    # renderizado de cada sección a PNG (sin mostrar la ventana).
    if ($env:HDZ_GUI_TEST_DIR) { $ui.txtCarpeta.Text = $env:HDZ_GUI_TEST_DIR }
    Lanzar-Escaneo
    # Gancho de test: simular la deselección de un archivo (nombre exacto)
    if ($env:HDZ_GUI_TEST_DESELECT) {
        $script:seleccion["$($env:HDZ_GUI_TEST_DESELECT)"] = $false
        Refrescar-Adaptacion
    }
    # Gancho de test: simular clic en la primera firma y reportar la selección + resaltado
    if ($env:HDZ_GUI_TEST_FIRMA -and @($script:firmaBordes).Count -gt 0) {
        $bf = $script:firmaBordes[0].Borde
        $ev = New-Object System.Windows.Input.MouseButtonEventArgs([System.Windows.Input.Mouse]::PrimaryDevice, 0, [System.Windows.Input.MouseButton]::Left)
        $ev.RoutedEvent = [System.Windows.UIElement]::MouseLeftButtonUpEvent
        $bf.RaiseEvent($ev)
        Write-Host "TEST_FIRMA: firmaSel = '$($script:firmaSel)' ; grosorBorde = $($script:firmaBordes[0].Borde.BorderThickness.Left)"
    }
    # Gancho de test: seleccionar torrent en modo PACK y comprobar el autorrelleno del nombre
    if ($env:HDZ_GUI_TEST_PACK) {
        Set-ChipValor "torrent" "PACK"
        Write-Host "TEST_PACK: nombrePack = '$($ui.txtPackNombre.Text)'"
    }
    # Gancho de test: poner descripción de ejemplo y activar la pestaña de vista previa
    if ($env:HDZ_GUI_TEST_PREVIEW) {
        $ui.upDescripcion.Text = "[center]`n[img=700]https://image.tmdb.org/t/p/original/logo.png[/img]`n[b]Friends[/b] — Temporada 1`n[/center]"
        $ui.tabPrevia.IsChecked = $true
        Write-Host "TEST_PREVIEW: panPreview visible = $($ui.panPreview.Visibility) ; tools = $($ui.panEdicionTools.Visibility) ; textbox = $($ui.upDescripcion.Visibility)"
    }
    # Gancho de test: flechas con UNA sola pestaña (deben quedar OCULTAS)
    if ($env:HDZ_GUI_TEST_FLECHAS1) {
        $ui.navSubida.IsChecked = $true
        $rt = $win.Content; $rt.Measure([System.Windows.Size]::new(1100,760)); $rt.Arrange([System.Windows.Rect]::new(0,0,1100,760)); $rt.UpdateLayout()
        Update-FlechasTabs
        Write-Host "TEST_FLECHAS1(1 tab): viewport=$([int]$ui.scrTabs.ViewportWidth) extent=$([int]$ui.scrTabs.ExtentWidth) flechas=$($ui.btnTabIzq.Visibility)"
    }
    # Gancho de test: menú lateral colapsado (solo iconos)
    if ($env:HDZ_GUI_TEST_COLAPSAR) {
        $script:sidebarColapsado = $true; Aplicar-Sidebar
        Write-Host "TEST_COLAPSAR: ancho sidebar=$($ui.sidebar.Width) navGeneral.Content='$($ui.navGeneral.Content)' lblSec=$($ui.lblSecMontaje.Visibility)"
    }
    # Gancho de test: varias pestañas con títulos largos (estilo MKVToolNix) para ver scroll/flechas
    if ($env:HDZ_GUI_TEST_TABSLONG) {
        $ui.upTitulo.Text = "Cómo Conocí a Vuestra Madre (2005) S01 1080p AMZN WEB-DL DD+ 5.1 H.264-HDZ"
        foreach ($tt in @("Friends (1994) S01 2160p HMAX WEB-DL DV HDR10-HDZ", "La Leyenda de Korra (2012) S01 1080p NF WEB-DL-HDZ")) {
            Nueva-Tab; $ui.upTitulo.Text = $tt
        }
        Construir-TabBar
        Write-Host "TEST_TABSLONG: $(@($script:tabs).Count) pestañas"
        # Comprobar flechas con estas 3 pestañas largas (deben SALIR por desbordamiento)
        $ui.navSubida.IsChecked = $true
        $rt = $win.Content; $rt.Measure([System.Windows.Size]::new(1100,760)); $rt.Arrange([System.Windows.Rect]::new(0,0,1100,760)); $rt.UpdateLayout()
        Update-FlechasTabs
        Write-Host "TEST_FLECHAS(3 tabs): viewport=$([int]$ui.scrTabs.ViewportWidth) extent=$([int]$ui.scrTabs.ExtentWidth) flechas=$($ui.btnTabIzq.Visibility)"
        # Simular arrastre del asa de la descripción (+150 px) y comprobar el alto
        $h0 = $ui.gridDesc.Height
        $ev = New-Object System.Windows.Controls.Primitives.DragDeltaEventArgs(0, 150)
        $ev.RoutedEvent = [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent
        $ui.gripDesc.RaiseEvent($ev)
        Write-Host "TEST_RESIZE: gridDesc alto inicial=$h0 -> tras +150 = $($ui.gridDesc.Height) (esperado 450)"
    }
    # Gancho de test: secciones por sub único ambiguo + mapa que se enviaría
    if ($env:HDZ_GUI_TEST_SUBSEC) {
        Construir-FilasSubsUnicos @(
            [PSCustomObject]@{ Cod="eng"; Fmt="Text"; Label="Inglés · Texto" },
            [PSCustomObject]@{ Cod="es";  Fmt="Text"; Label="Castellano · Texto" }
        )
        # Marcar el 2º como Forzado
        $script:subsUnicosRows[1].RbForzado.IsChecked = $true
        $mapa = [ordered]@{}
        foreach ($r in @($script:subsUnicosRows)) { $mapa["$($r.Cod)|$($r.Fmt)"] = $(if ($r.RbForzado.IsChecked) { "Forzado" } else { "Completo" }) }
        Write-Host "TEST_SUBSEC: filas=$(@($script:subsUnicosRows).Count) mapa=$(($mapa.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ')"
    }
    # Gancho de test: pestañas de subida (snapshot/restore de estado independiente)
    if ($env:HDZ_GUI_TEST_PROG) {
        $ui.panProgreso.Visibility = "Visible"; $ui.barProgreso.Value = 42; $ui.lblProgresoPct.Text = "42%"; $ui.lblProgreso.Text = "Procesando 2/24: Friends S02E02"
        $ui.panProgreso2.Visibility = "Visible"; $ui.barProgreso2.Value = 73; $ui.lblProgreso2Pct.Text = "73%"; $ui.lblProgreso2.Text = "Ensamblando Friends S02E02"
        Set-ContenidoConsola $true   # estado "ocultar" → glifo de ojo tachado
        Write-Host "TEST_PROG: dos barras visibles + boton consola en modo ocultar"
    }
    if ($env:HDZ_GUI_TEST_TABS) {
        $ui.upTitulo.Text = "Pelicula A"; $ui.upTmdb.Text = "111"
        Nueva-Tab
        $ui.upTitulo.Text = "Pelicula B"; $ui.upTmdb.Text = "222"
        $r1 = "tras crear 2 tabs: activa=$($script:tabActual) titulo='$($ui.upTitulo.Text)' tmdb='$($ui.upTmdb.Text)'"
        Cambiar-Tab 0
        $r2 = "vuelvo a tab 0: titulo='$($ui.upTitulo.Text)' tmdb='$($ui.upTmdb.Text)' (esperado Pelicula A / 111)"
        Cambiar-Tab 1
        $r3 = "voy a tab 1: titulo='$($ui.upTitulo.Text)' tmdb='$($ui.upTmdb.Text)' (esperado Pelicula B / 222)"
        Write-Host "TEST_TABS: nTabs=$(@($script:tabs).Count)`n  $r1`n  $r2`n  $r3"
        Cerrar-Tab 1
        Write-Host "TEST_TABS: tras cerrar tab1 -> nTabs=$(@($script:tabs).Count) activa=$($script:tabActual) titulo='$($ui.upTitulo.Text)'"
    }
    # Gancho de test: construir miniaturas de capturas desde una carpeta con .jpg
    if ($env:HDZ_GUI_TEST_CAPSDIR -and (Test-Path -LiteralPath $env:HDZ_GUI_TEST_CAPSDIR)) {
        $jpgs = @(Get-ChildItem -LiteralPath $env:HDZ_GUI_TEST_CAPSDIR -File -Filter *.jpg | Select-Object -First 4 -ExpandProperty FullName)
        if ($jpgs.Count -gt 0) { Construir-Capturas $jpgs; Write-Host "TEST_CAPS: $($jpgs.Count) miniatura(s)" }
    }
    # Gancho de test: construir logos de prueba, simular clic en el 2º y comprobar resaltado
    if ($env:HDZ_GUI_TEST_LOGO) {
        Construir-Logos @(
            [PSCustomObject]@{ Url = "https://x/logo1.png"; Width = 800; Height = 300; Lang = "es" },
            [PSCustomObject]@{ Url = "https://x/logo2.png"; Width = 600; Height = 200; Lang = "es" }
        )
        $b1 = $script:logoBordes[1].Borde
        $ev = New-Object System.Windows.Input.MouseButtonEventArgs([System.Windows.Input.Mouse]::PrimaryDevice, 0, [System.Windows.Input.MouseButton]::Left)
        $ev.RoutedEvent = [System.Windows.UIElement]::MouseLeftButtonUpEvent
        $b1.RaiseEvent($ev)
        Write-Host "TEST_LOGO: logoSel='$($script:logoSel)' grosor[1]=$($script:logoBordes[1].Borde.BorderThickness.Left) grosor[0]=$($script:logoBordes[0].Borde.BorderThickness.Left)"
    }
    # Gancho de test: modo heterogéneo con 3 películas → pestañas de proyecto por archivo
    if ($env:HDZ_GUI_TEST_HET) {
        $nM="The.Matrix.1999.2160p.NF.WEB-DL.mkv"; $nF="Friends.S01.1994.1080p.HMAX.WEB-DL.mkv"; $nI="Interstellar.2014.1080p.AMZN.WEB-DL.mkv"
        # Análisis simulado: Matrix con DTS + sub und; Friends con audio und + PGS; Interstellar 4 idiomas de subs.
        $script:ultimoScan = @(
            @{ Nombre=$nM; Ruta="C:\v\$nM"; Error=$null; Video=@{Codec="hevc";Height=2160;EsDV=$false;EsHDR=$true}
               Audios=@(@{Index=0;Codec="dts";Profile="MA";Lang="eng";Title="";Forced=$false;Channels=6}, @{Index=1;Codec="eac3";Profile="";Lang="spa";Title="";Forced=$false;Channels=6})
               Subs=@(@{Index=2;Codec="subrip";Lang="und";Title="";Forced=$false;EsPGS=$false}) }
            @{ Nombre=$nF; Ruta="C:\v\$nF"; Error=$null; Video=@{Codec="h264";Height=1080;EsDV=$false;EsHDR=$false}
               Audios=@(@{Index=0;Codec="aac";Profile="";Lang="und";Title="";Forced=$false;Channels=2})
               Subs=@(@{Index=1;Codec="hdmv_pgs_subtitle";Lang="spa";Title="";Forced=$false;EsPGS=$true}) }
            @{ Nombre=$nI; Ruta="C:\v\$nI"; Error=$null; Video=@{Codec="hevc";Height=1080;EsDV=$false;EsHDR=$false}
               Audios=@(@{Index=0;Codec="truehd";Profile="";Lang="eng";Title="";Forced=$false;Channels=8})
               Subs=@(@{Index=1;Codec="subrip";Lang="fre";Title="";Forced=$false;EsPGS=$false}, @{Index=2;Codec="subrip";Lang="ger";Title="";Forced=$false;EsPGS=$false}, @{Index=3;Codec="subrip";Lang="ita";Title="";Forced=$false;EsPGS=$false}, @{Index=4;Codec="subrip";Lang="spa";Title="";Forced=$false;EsPGS=$false}) }
        )
        $script:gruposProy = @(
            @{ Clave="The Matrix";  Principal=$nM; Archivos=@($nM) },
            @{ Clave="Friends";     Principal=$nF; Archivos=@($nF) },
            @{ Clave="Interstellar";Principal=$nI; Archivos=@($nI) }
        )
        Set-ChipValor "modoLote" "HETEROGENEO"
        Actualizar-ModoLote
        $rA = "tab Matrix (activa): cardConvAudio=$($ui.cardConvAudio.Visibility) nOpcsOrigen=$(@($script:convTrackOpts).Count) cardUndSub=$($ui.cardUndSub.Visibility) cardPGS=$($ui.cardPGS.Visibility)"
        # Añadir una conversión en Matrix, ir a Friends (debe verse PGS), volver y comprobar que se conserva.
        [void](Add-FilaConv)
        Cambiar-TabProy 1
        $rB = "tab Friends: cardConvAudio=$($ui.cardConvAudio.Visibility) nFilasConv=$(@($script:convRows).Count) cardPGS=$($ui.cardPGS.Visibility)"
        Cambiar-TabProy 2
        $rC = "tab Interstellar: nOpcsOrigen=$(@($script:convTrackOpts).Count) nFilasConv=$(@($script:convRows).Count) cardFiltro=$($ui.cardFiltro.Visibility)"
        Cambiar-TabProy 0
        $rD = "vuelvo a Matrix: nFilasConv=$(@($script:convRows).Count) (esperado 2: la base + la añadida)"
        # Gathering del config por archivo (replica la lógica de envío): conversiones y und por película.
        $script:tabsProy[$script:tabProyActual].Decisiones = Snapshot-DecisionesTab
        $resumen = @($script:tabsProy | ForEach-Object {
            $d = $_.Decisiones
            $nund = if ($d.Und) { @($d.Und.Keys).Count } else { 0 }
            "$($_.Principal.Substring(0,12))…: Conv=$(@($d.Conv).Count) undPistas=$nund"
        }) -join " | "
        Write-Host "TEST_HET_AV:`n  $rA`n  $rB`n  $rC`n  $rD`n  config: $resumen"
        $r0 = "tras reconstruir: nTabs=$(@($script:tabsProy).Count) activa=$($script:tabProyActual) titulo='$($ui.txtTitulo.Text)' año='$($ui.txtAno.Text)' caps=$(Get-ChipValor 'capturasProy')"
        # Editar plataforma + capturas de la peli 0, ir a la 1 (otras capturas) y volver: deben conservarse.
        $ui.txtPlataformaOtra.Text = "EDITADO-NF"; Set-ChipValor "capturasProy" 20
        Cambiar-TabProy 1
        Set-ChipValor "capturasProy" 3
        $r1 = "tab 1: titulo='$($ui.txtTitulo.Text)' platOtra='$($ui.txtPlataformaOtra.Text)' caps=$(Get-ChipValor 'capturasProy')"
        Cambiar-TabProy 0
        $r2 = "vuelvo a tab 0: titulo='$($ui.txtTitulo.Text)' platOtra='$($ui.txtPlataformaOtra.Text)' caps=$(Get-ChipValor 'capturasProy') (esperado EDITADO-NF / caps 20)"
        # Mapa ProyectoPorArchivo que se enviaría (incluye NumCapturas por archivo)
        if ($script:tabProyActual -ge 0) { $script:tabsProy[$script:tabProyActual].Estado = Snapshot-Proyecto }
        $r3 = @($script:tabsProy | ForEach-Object { "$($_.Principal.Substring(0,[Math]::Min(20,$_.Principal.Length)))…=caps:$($_.Estado.Capturas)" }) -join " | "
        Write-Host "TEST_HET:`n  $r0`n  $r1`n  $r2`n  caps por peli: $r3"
    }
    try {
        $w = 1100; $h = 760
        $root = $win.Content
        if ($env:HDZ_GUI_TEST_TMDBKEY) { $ui.cfgTmdbKey.Text = $env:HDZ_GUI_TEST_TMDBKEY }
        if ($env:HDZ_GUI_TEST_TORRENT) { Cargar-TorrentParaSubir $env:HDZ_GUI_TEST_TORRENT }
        foreach ($sec in @("navGeneral", "navProyecto", "navAudio", "navSubs", "navSubida", "navAjustes")) {
            $ui[$sec].IsChecked = $true
            $root.Measure([System.Windows.Size]::new($w, $h))
            $root.Arrange([System.Windows.Rect]::new(0, 0, $w, $h))
            $root.UpdateLayout()
            $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($w, $h, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32))
            $rtb.Render($root)
            $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
            $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
            $png = Join-Path $env:TEMP "hdz_gui_$sec.png"
            $fs = [System.IO.File]::Open($png, [System.IO.FileMode]::Create)
            $enc.Save($fs); $fs.Close()
            Write-Host "Captura: $png"
        }
        # Render ALTO de un panel concreto (para ver secciones largas sin scroll)
        if ($env:HDZ_GUI_TEST_TALL) {
            $navMap = @{ navSubida="panSubida"; navAjustes="panAjustes"; navGeneral="panGeneral"; navProyecto="panProyecto"; navAudio="panAudio"; navSubs="panSubs" }
            $ui[$env:HDZ_GUI_TEST_TALL].IsChecked = $true
            $pan = $ui[$navMap[$env:HDZ_GUI_TEST_TALL]]
            $hh = 2200
            $pan.Measure([System.Windows.Size]::new(820, $hh))
            $pan.Arrange([System.Windows.Rect]::new(0, 0, 820, $hh))
            $pan.UpdateLayout()
            $realH = [int][Math]::Min(2200, [Math]::Ceiling($pan.ActualHeight) + 20)
            $rtb2 = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(820, $realH, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32))
            $rtb2.Render($pan)
            $enc2 = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
            $enc2.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb2))
            $png2 = Join-Path $env:TEMP "hdz_gui_tall_$($env:HDZ_GUI_TEST_TALL).png"
            $fs2 = [System.IO.File]::Open($png2, [System.IO.FileMode]::Create)
            $enc2.Save($fs2); $fs2.Close()
            Write-Host "Captura alta: $png2"
        }
        Write-Host "HDZ-GUI: ventana construida y análisis ejecutado correctamente (modo test)."
    } catch {
        Write-Host "HDZ-GUI: error en modo test: $($_.Exception.Message)"
    }
} else {
    Lanzar-Escaneo
    [void]$win.ShowDialog()
}
