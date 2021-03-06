<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Log;

class AutoConvert extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'auto:convert';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Convert VB';

    protected $dirVBNET = "";

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        $this->dirVBNET = env('DIR_VBNET', "D:\\XAMPP\\Convert_VB\\src\\src_VB.NET");

        Log::debug("==============================Start auto convert==============================");

        $dirSource = $this->ask("From directory?");

        if (empty($dirSource)) {
            $this->error('From directory => required');
            return 0;
        }

        $dirVBNETTemp = $this->ask('To directory? (Default: '.$this->dirVBNET.')');

        if (!empty($dirVBNETTemp)) {
            $this->dirVBNET = $dirVBNETTemp;
        }

        $this->info('<fg=blue>Convert from directory: </>' . $dirSource);
        $this->info('<fg=blue>All file after auto convert will copy to: </>'. $this->dirVBNET);

        $files = collect(File::allFiles($dirSource, true));
        $fileVbproj = $files->filter(function($file) {
            return preg_match('/.*vbproj$/', $file->getFilename());
        })->first();

        if (empty($fileVbproj)) {
            $this->error("Can't find file *.vbproj");
            return 0;
        }

        $programId = explode('.', $fileVbproj->getFilename())[0];

        if (!File::copyDirectory($dirSource, $this->dirVBNET . DIRECTORY_SEPARATOR . $programId)) {
            $this->error("Can't copy to: " . $this->dirVBNET . DIRECTORY_SEPARATOR . $programId);
            return 0;
        }

        $newVBProjPath = $this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . $programId.".vbproj";

        // 2-4. [???????????????ID].vbproj.user??????????????????
        if (!File::delete($this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . $programId.".vbproj.user")) {
            // $this->warn("Can't delete file ".$programId.".vbproj.user");
        }

        $dirVBNETProject = $this->dirVBNET . DIRECTORY_SEPARATOR . $programId;
        $files = collect(File::allFiles($dirVBNETProject, true));
        $dirs = collect(File::directories($dirVBNETProject));

        foreach ($dirs as $dir) {
            File::deleteDirectory($dir);
        }
        
        foreach ($files as $file) {
            if (!preg_match('/'.$programId.'.*/', $file->getFilename()) || preg_match('/'.$programId.'_bas\.vb/', $file->getFilename()) || preg_match('/'.$programId.'\.log/', $file->getFilename()) || preg_match('/'.$programId.'\.ico/', $file->getFilename())) {
                File::delete($file->getPathname());
            }
        }

        $dirFileBas = $this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . 'Bas_'.$programId.'.vb';
        if (!File::copy(public_path('Bas_Template.vb'), $dirFileBas)) {
            $this->warn("Copy file Bas_ error");
        }

        $this->replaceInFileWithRegex('<PROGRAM_ID>', $programId, $dirFileBas, -1);
        // $formName = $this->ask('What is your form name?');
        // $this->replaceInFileWithRegex('<FORM_NAME>', mb_convert_encoding($formName, 'UTF-8'), $dirFileBas, -1);

        // Copy temp vbproj -> vbproj
        File::delete($newVBProjPath);
        File::copy(public_path('VBPROJ_TEMPLATE.vbproj'), $newVBProjPath);

        // Edit new file vbproj
        $this->replaceInFileWithRegex('<PROGRAM_ID>', $programId, $newVBProjPath, -1);
        
        // Free replace
        $this->info('Start auto convert...');
        $this->newLine();

        $files = collect(File::allFiles($dirVBNETProject, true));
        $arrFromToToolBarClick = [];
        $arrMnuFile = [];
        $arrMnuEdit = [];

        $bar = $this->output->createProgressBar($files->count());
        $bar->setFormat('File: %message%' . $this->createEnter() . ' %current%/%max% [%bar%] %percent:3s%%');
        $bar->start();
        
        foreach ($files as $file) {
            $matchesMainFilename = null;
            $mainFilename = '';

            $bar->setMessage($file->getFilename());

            if (preg_match('/(.*)\.Designer\.vb/', $file->getFilename(), $matchesMainFilename)) {
                $mainFilename = str_replace('_frm', '', $matchesMainFilename[1]);

                // For design
                File::replaceInFile('[Global]', 'Global', $file->getPathname());
                File::replaceInFile('[Partial]', 'Partial', $file->getPathname());
                File::replaceInFile('Public', 'Friend', $file->getPathname());
                File::replaceInFile('Friend Sub New()', 'Public Sub New()', $file->getPathname());
                $this->removeLineByKeySearch('\.OcxState', $file->getPathname(), true);

                File::replaceInFile('_mnu', 'mnu', $file->getPathname());
                File::replaceInFile('_opt', 'opt', $file->getPathname());
                File::replaceInFile('_cmd', 'cmd', $file->getPathname());
                File::replaceInFile('_xLabel', 'xLabel', $file->getPathname());
                File::replaceInFile('_lbl', 'lbl', $file->getPathname());
                File::replaceInFile('_PGrid', 'PGrid', $file->getPathname());
                File::replaceInFile('_chk', 'chk', $file->getPathname());

                $fileContent = File::get($file->getPathname());

                for ($idx = 0; $idx < 10; $idx++) { 
                    $matches = null;
                    if (preg_match('/_Toolbar1_Button'. $idx .'\.Name = "(.*)"/', $fileContent, $matches)) {
                        $arrFromToToolBarClick[$mainFilename]['_Toolbar1_Button'. $idx] = 'tb'.$matches[1];

                        File::replaceInFile('_Toolbar1_Button'. $idx, 'tb'.$matches[1], $file->getPathname());
                    }
                }

                $matches = null;
                $fileContent = File::get($file->getPathname());

                if (preg_match_all('/Friend WithEvents (mnuFILEItem\_\d*)/', $fileContent, $matches)) {
                    $arrMnuFile[$mainFilename] = $matches[1];
                }

                if (preg_match_all('/Friend WithEvents (mnuEDITItem\_\d*)/', $fileContent, $matches)) {
                    $arrMnuEdit[$mainFilename] = $matches[1];
                }

                $this->removeLineByKeySearch('Me\.Font', $file->getPathname(), true);
                $this->removeLineByKeySearch('.*\.SetIndex\(.*, CType\(.+, Short\)\)', $file->getPathname(), true);
                $this->removeLineByKeySearch('Me\.KeyPreview', $file->getPathname(), true);

                File::replaceInFile('AxxComboLib.AxxCombo', 'CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('AxxDropLib.AxxDrop', 'CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('AxxLabelLib.AxxLabel', 'System.Windows.Forms.Label', $file->getPathname());
                File::replaceInFile('AxxLabelLib.AxxLabelArray', 'System.Windows.Forms.Label', $file->getPathname());
                File::replaceInFile('AxxLabelArray', 'CoreLib.LabelS', $file->getPathname());
                File::replaceInFile('AxxCBtnLib.AxxCmdBtn', 'CoreLib.ButtonS', $file->getPathname());
                File::replaceInFile('AxxCmdBtnArray', 'CoreLib.ButtonS', $file->getPathname());
                File::replaceInFile('AxXOPTIONLib.AxxOption', 'System.Windows.Forms.RadioButton', $file->getPathname());
                File::replaceInFile('AxxOptionArray', 'System.Windows.Forms.RadioButton', $file->getPathname());
                File::replaceInFile('AxXCHECKLib.AxxCheck', 'CoreLib.CheckBoxV', $file->getPathname());
                File::replaceInFile('AxxCheckArray', 'CoreLib.CheckBoxV', $file->getPathname());
                File::replaceInFile('AxxListLib.AxxList', 'System.Windows.Forms.ListBox', $file->getPathname());
                File::replaceInFile('AxxTextLib.AxxText', 'CoreLib.UltraTextEditorC', $file->getPathname());
                File::replaceInFile('AxxDateLib.AxxDate', 'CoreLib.LabelS', $file->getPathname());
                File::replaceInFile('AxxDateLib.AxxTime', 'CoreLib.LabelS', $file->getPathname());
                File::replaceInFile('Microsoft.VisualBasic.Compatibility.VB6.Panel', 'System.Windows.Forms.Panel', $file->getPathname());
                File::replaceInFile('Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray', 'System.Windows.Forms.PictureBox', $file->getPathname());
                File::replaceInFile('AxPGRIDLib.AxPerfectGrid', 'CoreLib.UltraGridP', $file->getPathname());
                File::replaceInFile('AxxZipLib.AxxZip', 'CoreLib.LabelS', $file->getPathname());
                File::replaceInFile('AxxKanaLib.AxxKana', 'CoreLib.LabelS', $file->getPathname());
                File::replaceInFile('AxxNumLib.AxxNum', 'CoreLib.UltraNumericEditorE', $file->getPathname());

                $this->removeLineByKeySearch('CType\(Me\..*, System\.ComponentModel\.ISupportInitialize\)\.BeginInit\(\)', $file->getPathname(), true);
                $this->removeLineByKeySearch('CType\(Me\..*, System\.ComponentModel\.ISupportInitialize\)\.EndInit\(\)', $file->getPathname(), true);

                $this->removeLineByKeySearch('ImageList.+', $file->getPathname(), true);

                $this->removeLineByKeySearch('Microsoft\.VisualBasic\.Compatibility\.VB6\.ToolStripMenuItemArray', $file->getPathname(), true);

                // $this->removeLineByKeySearch('System\.Windows\.Forms\.ToolStripSeparator', $file->getPathname(), true);

                $this->removeLineByKeySearch('CrDraw1', $file->getPathname(), true);

                // Truong hop cu the (% cao la dung)
                File::replaceInFile('Friend WithEvents lblNMGB As System.Windows.Forms.Label', 'Friend WithEvents lblNMGB As CoreLib.LabelFaculty', $file->getPathname());
                File::replaceInFile('Me.lblNMGB = New System.Windows.Forms.Label', 'Me.lblNMGB = New CoreLib.LabelFaculty', $file->getPathname());

                File::replaceInFile('Friend WithEvents cboCDGK As CoreLib.ComboBoxL', 'Friend WithEvents cboCDGK As CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('Me.cboCDGK = New CoreLib.ComboBoxL', 'Me.cboCDGK = New CoreLib.UltraComboE', $file->getPathname());

                File::replaceInFile('Friend WithEvents xLabel2 As System.Windows.Forms.Label', 'Friend WithEvents xLabel2 As CoreLib.LabelS', $file->getPathname());
                File::replaceInFile('Me.xLabel2 = New System.Windows.Forms.Label', 'Me.xLabel2 = New CoreLib.LabelS', $file->getPathname());

                //File::replaceInFile('Friend WithEvents xLabel1 As System.Windows.Forms.Label', 'Friend WithEvents xLabel1 As CoreLib.LabelS', $file->getPathname());
                //File::replaceInFile('Me.xLabel1 = New System.Windows.Forms.Label', 'Me.xLabel1 = New CoreLib.LabelS', $file->getPathname());

                File::replaceInFile('Friend WithEvents cboSTCDGK As CoreLib.ComboBoxL', 'Friend WithEvents cboSTCDGK As CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('Me.cboSTCDGK = New CoreLib.ComboBoxL', 'Me.cboSTCDGK = New CoreLib.UltraComboE', $file->getPathname());

                File::replaceInFile('Friend WithEvents cboEDCDGK As CoreLib.ComboBoxL', 'Friend WithEvents cboEDCDGK As CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('Me.cboEDCDGK = New CoreLib.ComboBoxL', 'Me.cboEDCDGK = New CoreLib.UltraComboE', $file->getPathname());

                File::replaceInFile('Friend WithEvents cboSTKBSK As CoreLib.ComboBoxL', 'Friend WithEvents cboSTKBSK As CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('Me.cboSTKBSK = New CoreLib.ComboBoxL', 'Me.cboSTKBSK = New CoreLib.UltraComboE', $file->getPathname());

                File::replaceInFile('Friend WithEvents cboEDKBSK As CoreLib.ComboBoxL', 'Friend WithEvents cboEDKBSK As CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('Me.cboEDKBSK = New CoreLib.ComboBoxL', 'Me.cboEDKBSK = New CoreLib.UltraComboE', $file->getPathname());

                File::replaceInFile('????', '?t?@?C??', $file->getPathname());
                File::replaceInFile('????', '?w???v', $file->getPathname());

                $this->changeSizeToolBarButton($file->getPathname());

            }
            elseif (preg_match('/^('.$programId.'.*)\.vb/', $file->getFilename(), $matchesMainFilename)) {
                $mainFilename = str_replace('_frm', '', $matchesMainFilename[1]);

                // For logic file
                $this->replaceInFileWithRegex('System.Windows.Forms.Form', 'Frm_Core', $file->getPathname());

                $this->removeLineByKeySearch('UPGRADE\_', $file->getPathname(), true);

                $arrFileContent = file($file->getPathname());

                foreach ($arrFileContent as $key => $content) {
                    if (preg_match('/Option Explicit On/', $content)) {
                        $fromText = '';
                        $toText = '';

                        if (preg_match('/'.preg_quote('Imports VB = Microsoft.VisualBasic', '/').'/', $arrFileContent[$key + 1])) {
                            $fromText = $arrFileContent[$key + 1];
                            $toText = $arrFileContent[$key + 1] . 'Imports CoreLib' . $this->createEnter() . 'Imports CoreNS' . $this->createEnter();
                        } else {
                            $fromText = $arrFileContent[$key];
                            $toText = $arrFileContent[$key] . 'Imports VB = Microsoft.VisualBasic' . $this->createEnter() . 'Imports CoreLib' . $this->createEnter() . 'Imports CoreNS' . $this->createEnter();
                        }

                        if (preg_match('/mPRNDevice/', file_get_contents($file->getPathname()))) {
                            $toText = $toText . 'Imports CoReportsCoreU' . $this->createEnter() . 'Imports CoReportsU' . $this->createEnter(2);

                            // Add reference to vbproj
                            $this->replaceInFileWithRegex('</Reference>', '</Reference>' . $this->createEnter() .  $this->createTab() . '<Reference Include="Interop.CoReportsCoreU">' . $this->createEnter() . $this->createTab(2) . '<HintPath>..\..\DLL\Interop.CoReportsCoreU.dll</HintPath>' . $this->createEnter() . $this->createTab() . '</Reference>' . $this->createEnter() . $this->createTab() . '<Reference Include="Interop.CoReportsU">' . $this->createEnter() . $this->createTab(2) . '<HintPath>..\..\DLL\Interop.CoReportsU.dll</HintPath>' . $this->createEnter() . $this->createTab() . '</Reference>' .$this->createEnter(), $newVBProjPath);
                        
                            // Add link to file crf
                            // $this->replaceInFileWithRegex('</ItemGroup>', '</ItemGroup>' .$this->createEnter() . $this->createTab() . '<ItemGroup>' . $this->createEnter() . $this->createTab(2) . str_replace('<PROGRAM_ID>', $programId, '<None Include="..\..\Report\<PROGRAM_ID>.crf" Link="Report\<PROGRAM_ID>.crf">') . $this->createEnter() . $this->createTab(3) . '<CopyToOutputDirectory>Always</CopyToOutputDirectory>' . $this->createEnter() .$this->createTab(2) . '</None>' .$this->createEnter() . $this->createTab() . '</ItemGroup>' . $this->createEnter(), $newVBProjPath);

                        }

                        if (preg_match('/\?/', file_get_contents($file->getPathname()))) {
                            $toText = $toText . 'Imports System.Data' . $this->createEnter(2);
                        }

                        $this->replaceInFileWithRegex($fromText, $toText, $file->getPathname());

                        break;
                    }
                }

                File::replaceInFile('GoSub', 'GoTo', $file->getPathname());

                $arrTBName = ['PRINT', 'PREVIEW', 'CANCEL', 'EXIT', 'EXEC', 'ROWDELETE', 'COPY', 'ROWINSERT', 'EXCEL', 'DELETE', 'EDIT', 'HELP', 'PRE', 'AUTO'];
                foreach ($arrTBName as $tbName) {
                    File::replaceInFile('Toolbar1.Items.Item("'.$tbName.'").Enabled', 'tb'.$tbName.'.Enabled', $file->getPathname());
                    File::replaceInFile('Case "'.$tbName.'"', 'Case "tb'.$tbName.'"', $file->getPathname());
                }

                for ($idx = 0; $idx < 10; $idx++) {
                    File::replaceInFile('mnuFILEItem('.$idx.')', 'mnuFILEItem_'.$idx, $file->getPathname());
                    File::replaceInFile('mnuEDITItem('.$idx.')', 'mnuEDITItem_'.$idx, $file->getPathname());

                    File::replaceInFile('mnuFILEItem.Item('.$idx.')', 'mnuFILEItem_'.$idx, $file->getPathname());
                    File::replaceInFile('mnuEDITItem.Item('.$idx.')', 'mnuEDITItem_'.$idx, $file->getPathname());
                }

                $this->removeLineByKeySearch('BEGIN TRAN', $file->getPathname(), true, 'dbCon2.BeginTransaction()' . $this->createEnter());
                $this->removeLineByKeySearch('COMMIT TRAN', $file->getPathname(), true, 'dbCon2.Commit()' . $this->createEnter());
                $this->removeLineByKeySearch('ROLLBACK TRAN', $file->getPathname(), true, 'dbCon2.Rollback()' . $this->createEnter());

                $this->removeLineByKeySearch(preg_quote('Dim Index As Short =', '/'), $file->getPathname(), true, $this->createTab(2) . 'Dim Index As Short = FormUtil.getControlPosition(eventSender)' . $this->createEnter());

                // File::replaceInFile('CellCheck_Numeric(PGrid, ', 'CellCheck_Numeric(', $file->getPathname()); //Sai khi co cac man nhieu grid tren 1 man @@

                File::replaceInFile('VB6.Format(', 'Format(', $file->getPathname());
                File::replaceInFile('VB6.TwipsToPixelsY(', 'TwipsToPixelsY(Me, ', $file->getPathname());
                File::replaceInFile('VB6.TwipsToPixelsX(', 'TwipsToPixelsX(Me, ', $file->getPathname());
                File::replaceInFile('VB6.PixelsToTwipsX(', 'PixelsToTwipsX(Me, ', $file->getPathname());
                File::replaceInFile('VB6.PixelsToTwipsY(', 'PixelsToTwipsY(Me, ', $file->getPathname());
                File::replaceInFile('VB6.TwipsPerPixelX', 'TwipsPerPixelX(Me)', $file->getPathname());
                File::replaceInFile('yyyy/mm/dd', 'yyyy/MM/dd', $file->getPathname());
                File::replaceInFile('yyyymmdd', 'yyyyMMdd', $file->getPathname());
                File::replaceInFile('yyyy?Nmm??dd??', 'yyyy?NMM??dd??', $file->getPathname());
                File::replaceInFile('hh:nn', 'HH:mm', $file->getPathname());
                File::replaceInFile('hh:MM:ss', 'HH:mm:ss', $file->getPathname());

                File::replaceInFile('CrDraw1', 'mCrDraw', $file->getPathname());

                File::replaceInFile('.RecordCount', '.F ields("rCount").Value \'"ADODB.Recordset.recordCountSQLstrings"', $file->getPathname());

                File::replaceInFile('.set_CellEnabled(', '.set_CellEnable(', $file->getPathname());

                $this->replaceQestionMarkToText($file->getPathname());

                File::replaceInFile('Private mCrForm As CoReports.CrForm', 'Private mCrForm As CrForm' . $this->createEnter() . $this->createTab() . 'Private mCrDraw As CrDraw' . $this->createEnter() .$this->createTab() . 'Private Const mPaperSize As String = "A4"' . $this->createEnter(), $file->getPathname());
                File::replaceInFile('If pFncVal <> 0 Then', 'If FormUtil.isPrtEndError(pFncVal) Then', $file->getPathname());

                File::replaceInFile('AxPGRIDLib.AxPerfectGrid', 'CoreLib.UltraGridP', $file->getPathname());
                File::replaceInFile('AxxCBtnLib.AxxCmdBtn', 'CoreLib.ButtonS', $file->getPathname());
                File::replaceInFile('AxxLabelLib.AxxLabel', 'System.Windows.Forms.Label', $file->getPathname());
                File::replaceInFile('AxxComboLib.AxxCombo', 'CoreLib.UltraComboE', $file->getPathname());

                $this->replaceInFileWithRegex('pBytes = LenB(StrConv(pPGrid.get_CellText(Row, Col), vbFromUnicode))', 'Dim sutil As StringUtil = New StringUtil(StringUtil.ENC_SHIFTJIS)' . $this->createEnter() . $this->createTab(2) . 'pBytes = sutil.getByteCount(pPGrid.get_CellText(Row, Col))' . $this->createEnter(), $file->getPathname());

                $this->appendTextToFunction('frm'.$mainFilename.'_Load', $this->createTab(2) .'ImageListUtil.setToolStripImage(Toolbar1)', $file->getPathname(), 'mMBOXTitle = Me.Text');

                if (preg_match('/mPRNDevice/', file_get_contents($file->getPathname()))) {
                    $this->appendTextToFunction('frm'.$mainFilename.'_Load', $this->createTab(2) .'mCrDraw = New CrDraw()', $file->getPathname(), 'mMBOXTitle = Me.Text');
                }

                File::replaceInFile('ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs', 'ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs', $file->getPathname());
                File::replaceInFile('Handles Me.FormClosed', 'Handles Me.FormClosing', $file->getPathname());

                if (!preg_match('/'.preg_quote('If mnuFILEItem_9.Enabled = False Then', '/').'/', file_get_contents($file->getPathname()))) {
                    $this->appendTextToFunction('frm'.$mainFilename.'_FormClosed', $this->createTab(2) . 'If mnuFILEItem_9.Enabled = False Then' . $this->createEnter() . $this->createTab(3) . 'mMsgText = "?o?^???????????B?I?????????????B"' . $this->createEnter() . $this->createTab(3) . 'MsgBox(mMsgText, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation,mMBOXTitle)' . $this->createEnter() . $this->createTab(3) . 'eventArgs.Cancel = True' . $this->createEnter() . $this->createTab(3) . 'Exit Sub' . $this->createEnter() . $this->createTab(2) . 'End if' . $this->createEnter(), $file->getPathname(), 'Sub frm'.$mainFilename.'_FormClosed');
                }

                // Add dbCmd, dbRec -> Bas_
                $matchesDbCmd = null;
                $matchesDbRec = null;

                if (preg_match_all('/^\s*(dbCmd\S*?)\./im', file_get_contents($file->getPathname()), $matchesDbCmd)) {
                    $matchesDbCmd = array_unique($matchesDbCmd[1]);
                    
                    foreach ($matchesDbCmd as $matchDbCmd) {
                        if (!preg_match('/' . $matchDbCmd . '/', file_get_contents($dirFileBas))) {
                            $this->removeLineByKeySearch(preg_quote('Public dbCmdUPD As CoreLib.ADODB.Command', '/'), $dirFileBas, true, $this->createTab().'Public dbCmdUPD As CoreLib.ADODB.Command' . $this->createEnter() . $this->createTab(). 'Public '.$matchDbCmd.' As CoreLib.ADODB.Command' . $this->createEnter());
                            $this->removeLineByKeySearch(preg_quote('If Not dbCmdUPD Is Nothing Then dbCmdUPD.Dispose()', '/'), $dirFileBas, true, $this->createTab(3).'If Not dbCmdUPD Is Nothing Then dbCmdUPD.Dispose()' . $this->createEnter() . $this->createTab(3). 'If Not '.$matchDbCmd.' Is Nothing Then '.$matchDbCmd.'.Dispose()' . $this->createEnter());
                            $this->removeLineByKeySearch(preg_quote('dbCmdUPD = Nothing', '/'), $dirFileBas, true, $this->createTab(3).'dbCmdUPD = Nothing' . $this->createEnter() . $this->createTab(3). $matchDbCmd . ' = Nothing' . $this->createEnter());
                        }

                        $fileContent = File::get($file->getPathname());
                        for ($idx = 0; $idx < 100; $idx++) {
                            if (preg_match('/'.$matchDbCmd.'\.Parameters\('.$idx.'\)\.Value = System\.DBNull\.Value/', $fileContent)) {
                                File::replaceInFile($matchDbCmd.'.Parameters('.$idx.')', $matchDbCmd.'.Parameters.Add("@p'.$idx.'", SqlDbType.NText)', $file->getPathname());
                            } else {
                                File::replaceInFile($matchDbCmd.'.Parameters('.$idx.')', $matchDbCmd.'.Parameters.Add("@p'.$idx.'", SqlDbType.Int)', $file->getPathname());
                            }
                        }
                    }
                }

                if (preg_match_all('/^\s*(dbRec\S*?)\./im', file_get_contents($file->getPathname()), $matchesDbRec)) {
                    $matchesDbRec = array_unique($matchesDbRec[1]);
                    
                    foreach ($matchesDbRec as $matchDbRec) {
                        if (!preg_match('/' . $matchDbRec . '/', file_get_contents($dirFileBas))) {
                            $this->removeLineByKeySearch(preg_quote('Public dbCmdUPD As CoreLib.ADODB.Command', '/'), $dirFileBas, true, $this->createTab().'Public dbCmdUPD As CoreLib.ADODB.Command' . $this->createEnter() . $this->createTab(). 'Public '.$matchDbRec.' As CoreLib.ADODB.Recordset' . $this->createEnter());
                            $this->removeLineByKeySearch(preg_quote('If Not dbCmdUPD Is Nothing Then dbCmdUPD.Dispose()', '/'), $dirFileBas, true, $this->createTab(3).'If Not dbCmdUPD Is Nothing Then dbCmdUPD.Dispose()' . $this->createEnter() . $this->createTab(3). 'If Not '.$matchDbRec.' Is Nothing Then '.$matchDbRec.'.Close()' . $this->createEnter());
                            $this->removeLineByKeySearch(preg_quote('dbCmdUPD = Nothing', '/'), $dirFileBas, true, $this->createTab(3).'dbCmdUPD = Nothing' . $this->createEnter() . $this->createTab(3). $matchDbRec . ' = Nothing' . $this->createEnter());
                        }
                    }
                }

                $this->replaceFunctionToText('mnuHELPItem_Click', $this->createTab() . 'Public Sub mnuHELPItem_Click(ByVal eventSender As System.Object, ByVal EventArgs As System.EventArgs) Handles mnuHELPItem_0.Click' . $this->createEnter() . $this->createTab(2) . 'Dim Index As Short = FormUtil.getControlPosition(eventSender)' . $this->createEnter(2) . $this->createTab(2) . 'Select Case Index' . $this->createEnter() . $this->createTab(3) . 'Case 0' . $this->createEnter() .$this->createTab(4) . 'Using helpForm As New Frm_HelpScreen()' . $this->createEnter() . $this->createTab(5) . 'With helpForm' . $this->createEnter() . $this->createTab(6) . '.OptionMode = True' . $this->createEnter() . $this->createTab(6) . '.HelpFileName = My.Application.Info.AssemblyName' . $this->createEnter() . $this->createTab(6) . '.ShowDialog()' .$this->createEnter() .$this->createTab(5) . 'End With' . $this->createEnter() . $this->createTab(4) . 'End Using' . $this->createEnter() . $this->createTab(2) . 'End Select' . $this->createEnter(2) . $this->createTab() . 'End Sub' . $this->createEnter(), $file->getPathname());

                if (isset($arrFromToToolBarClick[$mainFilename]) && !empty($arrFromToToolBarClick[$mainFilename])) {
                    foreach ($arrFromToToolBarClick[$mainFilename] as $from => $to) {
                        File::replaceInFile($from, $to, $file->getPathname());
                    }
                }

                if (isset($arrMnuFile[$mainFilename]) && !empty($arrMnuFile[$mainFilename])) {
                    $arrMnuFile[$mainFilename] = array_map(function($item) {
                        return $item . '.Click';
                    }, $arrMnuFile[$mainFilename]);

                    $mnuFileString = join(', ', $arrMnuFile[$mainFilename]);

                    File::replaceInFile('Public Sub mnuFILEItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFILEItem.Click', 'Public Sub mnuFILEItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ' . $mnuFileString, $file->getPathname());
                }

                if (isset($arrMnuEdit[$mainFilename]) && !empty($arrMnuEdit[$mainFilename])) {
                    $arrMnuEdit[$mainFilename] = array_map(function($item) {
                        return $item . '.Click';
                    }, $arrMnuEdit[$mainFilename]);

                    $mnuEditString = join(', ', $arrMnuEdit[$mainFilename]);

                    File::replaceInFile('Public Sub mnuEDITItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEDITItem.Click', 'Public Sub mnuEDITItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ' . $mnuEditString, $file->getPathname());
                }

                File::replaceInFile('.get_ColStyle(pCOL)', '.DisplayLayout.Bands(0).Columns(pCOL).Style', $file->getPathname());

                File::replaceInFile('_CellGotFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellGotFocusEvent)', '_CellGotFocus(ByVal eventSender As System.Object, ByVal eventArgs As EventArgs)', $file->getPathname());
                File::replaceInFile('.CellGotFocus', '.AfterCellActivate', $file->getPathname());

                File::replaceInFile('_CellLostFocus(ByVal eventSender As System.Object, ByVal eventArgs As AxPGRIDLib._DPGridEvents_CellLostFocusEvent)', '_CellLostFocus(ByVal eventSender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs)', $file->getPathname());
                File::replaceInFile('.CellLostFocus', '.BeforeCellActivate', $file->getPathname());

                File::replaceInFile('dbCon2.Execute(SqlText,  , ADODB.CommandTypeEnum.adCmdText + ADODB.ExecuteOptionEnum.adExecuteNoRecords)', 'dbCon2.Execute(SqlText, pRecordsAffected, ADODB.CommandTypeEnum.adCmdText + ADODB.ExecuteOptionEnum.adExecuteNoRecords)', $file->getPathname());

                $this->replaceFunctionToText('cboPRN_SelectedIndexChanged', '', $file->getPathname());
                $this->replaceFunctionToText('cboPRN_KeyDown', '', $file->getPathname());
                $this->replaceFunctionToText('cboPRN_KeyPress', '', $file->getPathname());
                $this->removeLineByKeySearch('Dim Printer As New Printer', $file->getPathname(), true);
                $this->removeLineByKeySearch(preg_quote('mPRNDevice = Printer.DeviceName', '/'), $file->getPathname(), true);
                $this->removeLineByKeySearch(preg_quote('mPRNDriver = Printer.DriverName', '/'), $file->getPathname(), true);
                $this->removeLineByKeySearch(preg_quote('mPRNPort = Printer.Port', '/'), $file->getPathname(), true);
                $this->removeLineByKeySearch(preg_quote('lblPaperSize.Text = "?`?S?F?c"', '/'), $file->getPathname(), true);
                $this->removeLineByKeySearch(preg_quote('Call ComboPrnSet(Me, mPRNDevice)', '/'), $file->getPathname(), true);

                File::replaceInFile('Format(pYMD, "yyyy?N?x")', 'Format(Convert.ToDateTime(pYMD), "yyyy?N?x")', $file->getPathname());
                File::replaceInFile('????????', 'center', $file->getPathname());
                File::replaceInFile('??????', 'left', $file->getPathname());
                File::replaceInFile('?E????', 'right', $file->getPathname());

                $this->removeLineByKeySearch('Private ExcelApp As New Microsoft\.Office\.Interop\.Excel\.Application', $file->getPathname(), true);

                File::replaceInFile('.PaperSize = CoReportsCore.corPaperSize.corPaperA4', '.PaperSize = corPaperSize.corPaperA4', $file->getPathname());
                File::replaceInFile('.ObjectType = CoReports.corObjectType.corList', '.ObjectType = corObjectType.corList', $file->getPathname());

                $this->removeLineByKeySearch(preg_quote("' ????", '/'), $file->getPathname(), true);

            }

            $bar->advance();
        }

        $bar->finish();

        $this->newLine(2);
        $this->info('<fg=green>Auto convert success...</>');

        return 0;
    }

    protected function createTab($num = 1) {
        return str_repeat(chr(9), $num);
    }

    protected function createEnter($num = 1) {
        return str_repeat(PHP_EOL, $num);
    }

    protected function removeLineByKeySearch($keySearch, $path, $isRegex, $toString = '') {
        $arrFileContent = file($path);
        $exceptText = ['UPGRADE\_', preg_quote("' ????", '/')];

        foreach ($arrFileContent as $content) {
            if (preg_match('/^\s*\'.*/', $content) && !in_array($keySearch, $exceptText)) {
                continue;
            }

            if ($isRegex) {
                if (preg_match('/.*'.$keySearch.'.*/', $content)) { //cho nay nen bo regex 2 ben mac dinh di
                    File::replaceInFile($content, $toString, $path);
                }
            } else {
                // without regex

            }
        }
    }

    protected function changeSizeToolBarButton($path) {
        $arrFileContent = file($path);
        $matchTbName = null;

        foreach ($arrFileContent as $fileContent) {
            if (preg_match('/^\s*\'.*/', $fileContent)) {
                continue;
            }

            if (preg_match('/Me\.(tb.*)\.Size \= New System\.Drawing\.Size\(\d\d\, \d\d\)/', $fileContent, $matchTbName)) {
                $this->replaceInFileWithRegex($fileContent, $this->createTab(2) . 'Me.'.$matchTbName[1].'.Size = New System.Drawing.Size(56, 47)' . $this->createEnter(), $path);
            }
        }
    }

    protected function replaceQestionMarkToText($path) {
        $arrFileContent = file($path);

        $start = 0;
        $arrCheckToResetStart = ['delete', 'select', 'insert', 'update'];
        $countQuestionMark = substr_count(file_get_contents($path), '?');

        foreach ($arrFileContent as $fileContent) {
            if (preg_match('/^\s*\'.*/', $fileContent)) {
                continue;
            }

            $newFileContent = $fileContent;

            foreach ($arrCheckToResetStart as $itemCheckToResetStart) {
                if (strpos($fileContent, $itemCheckToResetStart) !== false) {
                    $start = 0;
                }
            }

            while (strpos($newFileContent, '?') !== false) {
                $newFileContent = preg_replace('/\?/', '@p'.$start, $newFileContent, 1);

                $start++;
                $countQuestionMark--;
            }

            if ($fileContent != $newFileContent) {
                $this->replaceInFileWithRegex($fileContent, $newFileContent, $path);
            }

            if ($countQuestionMark <= 0) {
                break;
            }
        }
    }

    protected function replaceInFileWithRegex($search, $replace, $path, $limit = 1)
    {
        file_put_contents($path, preg_replace('/'.preg_quote($search, '/').'/', $replace, file_get_contents($path), $limit));
    }

    protected function appendTextToFunction($nameFunc, $textAppend, $path, $afterText = null) {
        $arrFileContent = file($path);
        $startFunc = false;

        foreach ($arrFileContent as $key => $content) {
            if (preg_match('/^\s*\'.*/', $content)) {
                continue;
            }

            if (preg_match('/'.preg_quote('Sub ' . $nameFunc, '/').'/', $content)) {
                $startFunc = true;
            }

            if ($startFunc) {
                if ($afterText !== null) {
                    if (preg_match('/'.preg_quote($afterText, '/').'/', $content)) {
                        $this->replaceInFileWithRegex($arrFileContent[$key - 2] . $arrFileContent[$key - 1] . $arrFileContent[$key], $arrFileContent[$key - 2] . $arrFileContent[$key - 1] . $arrFileContent[$key] . $textAppend . $this->createEnter(), $path);
                        $startFunc = false;
                    }
                } else {
                    // End func : chua xu li
                    
                }

                if (preg_match('/End Sub/', $content)) {
                    $startFunc = false;
                    break;
                }
            }
        }
    }

    protected function replaceFunctionToText($nameFunc, $textReplace, $path) {
        $arrFileContent = file($path);
        $startFunc = false;
        $fromText = '';

        foreach ($arrFileContent as $key => $content) {
            if (preg_match('/^\s*\'.*/', $content)) {
                continue;
            }
            
            if (preg_match('/'.preg_quote('Sub ' . $nameFunc, '/').'/', $content)) {
                $startFunc = true;
            }

            if ($startFunc) {
                $fromText = $fromText . $content;

                if (preg_match('/End Sub/', $content)) {
                    $startFunc = false;
                    break;
                }
            }
        }

        if ($fromText !== '') {
            $this->replaceInFileWithRegex($fromText, $textReplace, $path);
        }
    }
}
