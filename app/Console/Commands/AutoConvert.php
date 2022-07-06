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
    protected $signature = 'auto:convert {dirSource}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Convert VB';

    protected $dirDevenv = "";
    protected $dirVBNET = "";

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        $this->dirDevenv = env('DIR_DEVENV', "C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\Common7\\IDE\\devenv.exe");
        $this->dirVBNET = env('DIR_VBNET', "D:\\XAMPP\\Convert_VB\\src\\src_VB.NET");

        Log::debug("==============================Start auto convert==============================");

        $dirSource = $this->argument('dirSource');

        Log::debug("Convert from dir: " . $dirSource);

        Log::debug("Search file *.vbproj...");
        $files = collect(File::allFiles($dirSource, true));
        $fileVbproj = $files->filter(function($file) {
            return preg_match('/.*vbproj$/', $file->getFilename());
        })->first();

        if (empty($fileVbproj)) {
            Log::debug("Can't find file *.vbproj");
            return 0;
        }

        Log::debug("Find file *.vbproj success...");
        $programId = explode('.', $fileVbproj->getFilename())[0];

        Log::debug("Start VBUpgrade...");
        Log::debug("Using ".$this->dirDevenv." ...");
        // 2-1
        $output = null;
        // dd('"' . $this->dirDevenv . '"' . ' ' . '"' . $fileVbproj->getPathname() . '"' . ' /upgrade', $output);
        if (false === exec('"' . $this->dirDevenv . '"' . ' ' . '"' . $fileVbproj->getPathname() . '"' . ' /upgrade', $output)) {
            Log::debug("Can't VBUpgrade...");
            return 0;
        }

        Log::debug("VBUpgrade success...");
        Log::debug("Start update net framework...");
        Log::debug("Using upgrade-asssistant...");

        if (false === exec('upgrade-assistant upgrade "' . $fileVbproj->getPathname() . '" --non-interactive', $output)) {
            Log::debug("Can't update net framework...");
            return 0;
        }

        Log::debug("Update net framework success...");

        Log::debug("Copy dir from " . $dirSource);
        Log::debug("To " . $this->dirVBNET . DIRECTORY_SEPARATOR . $programId);
        
        if (!File::copyDirectory($dirSource, $this->dirVBNET . DIRECTORY_SEPARATOR . $programId)) {
            Log::debug("Can't copy dir...");
            return 0;
        }

        Log::debug("Copy dir success...");

        $newVBProjPath = $this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . $programId.".vbproj";

        // 2-4. [ãƒ—ãƒ­ã‚°ãƒ©ãƒ?ID].vbproj.userã‚’å‰Šé™¤ã™ã‚‹ã€?
        Log::debug("Delete file ".$programId.".vbproj.user");
        
        if (!File::delete($this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . $programId.".vbproj.user")) {
            Log::debug("Can't delete file ".$programId.".vbproj.user");
            Log::debug("Continu.......");
        }

        Log::debug("Delete file success...");

        Log::debug("Start delete other files...");
        $dirVBNETProject = $this->dirVBNET . DIRECTORY_SEPARATOR . $programId;
        $files = collect(File::allFiles($dirVBNETProject, true));
        $dirs = collect(File::directories($dirVBNETProject));

        foreach ($dirs as $dir) {
            File::deleteDirectory($dir);
        }
        
        foreach ($files as $file) {
            if (!preg_match('/'.$programId.'.*/', $file->getFilename()) || preg_match('/'.$programId.'_bas\.vb/', $file->getFilename()) || preg_match('/'.$programId.'\.log/', $file->getFilename())) {
                File::delete($file->getPathname());
            }
        }

        Log::debug("Copy file Bas_");
        if (!File::copy($this->dirVBNET . DIRECTORY_SEPARATOR . 'Bas_Template.vb', $this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . 'Bas_'.$programId.'.vb')) {
            Log::debug("Copy file Bas_ error");
        }

        // Log::debug("Edit file Bas_");
        // File::replaceInFile(['[ãƒ—ãƒ­ã‚°ãƒ©ãƒ?ID]'], $programId, $this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . 'Bas_'.$programId.'.vb');
        
        // Free replace
        $files = collect(File::allFiles($dirVBNETProject, true));

        foreach ($files as $file) {
            if (preg_match('/.*\.Designer\.vb/', $file->getFilename())) {
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

                $fileContent = File::get($file->getPathname());
                for ($idx = 0; $idx < 10; $idx++) { 
                    $matches = null;
                    if (preg_match('/_Toolbar1_Button'. $idx .'\.Name = "(.*)"/', $fileContent, $matches)) {
                        File::replaceInFile('_Toolbar1_Button'. $idx, 'tb'.$matches[1], $file->getPathname());
                    }
                }

                $this->removeLineByKeySearch('Me\.Font', $file->getPathname(), true);
                $this->removeLineByKeySearch('.*\.SetIndex\(.*, CType\(.+, Short\)\)', $file->getPathname(), true);
                $this->removeLineByKeySearch('Me\.KeyPreview', $file->getPathname(), true);

                File::replaceInFile('AxxComboLib.AxxCombo', 'CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('AxxDropLib.AxxDrop', 'CoreLib.ComboBoxL', $file->getPathname());
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

                File::replaceInFile('Friend WithEvents lblNMGB As System.Windows.Forms.Label', 'Friend WithEvents lblNMGB As CoreLib.LabelFaculty', $file->getPathname());
                File::replaceInFile('Me.lblNMGB = New System.Windows.Forms.Label', 'Me.lblNMGB = New CoreLib.LabelFaculty', $file->getPathname());

                File::replaceInFile('Friend WithEvents cboCDGK As CoreLib.ComboBoxL', 'Friend WithEvents cboCDGK As CoreLib.UltraComboE', $file->getPathname());
                File::replaceInFile('Me.cboCDGK = New CoreLib.ComboBoxL', 'Me.cboCDGK = New CoreLib.UltraComboE', $file->getPathname()); // Check lai thang nay dang chay khong dung


            }
            elseif (preg_match('/^'.$programId.'.*\.vb/', $file->getFilename())) {
                // For logic file
                $arrFileContent = file($file->getPathname());

                foreach ($arrFileContent as $key => $content) {
                    if (preg_match('/Option Explicit On/', $content)) {
                        $fromText = '';
                        $toText = '';

                        if (preg_match('/'.preg_quote('Imports VB = Microsoft.VisualBasic', '/').'/', $arrFileContent[$key + 1])) {
                            $fromText = $arrFileContent[$key + 1];
                            $toText = $arrFileContent[$key + 1] . 'Imports CoreLib' . PHP_EOL . 'Imports CoreNS' . PHP_EOL;
                        } else {
                            $fromText = $arrFileContent[$key];
                            $toText = $arrFileContent[$key] . 'Imports VB = Microsoft.VisualBasic' . PHP_EOL . 'Imports CoreLib' . PHP_EOL . 'Imports CoreNS' . PHP_EOL;
                        }

                        if (preg_match('/mPRNDevice/', file_get_contents($file->getPathname()))) {
                            $toText = $toText . 'Imports CoReportsCoreU' . PHP_EOL . 'Imports CoReportsU' . PHP_EOL . PHP_EOL;
                        }

                        $this->replaceInFileWithRegex($fromText, $toText, $file->getPathname());

                        break;
                    }
                }

                $this->replaceInFileWithRegex('System.Windows.Forms.Form', 'Frm_Core', $file->getPathname());

                $this->removeLineByKeySearch('UPGRADE_ISSUE', $file->getPathname(), true);
                $this->removeLineByKeySearch('UPGRADE_WARNING', $file->getPathname(), true);
                $this->removeLineByKeySearch('UPGRADE_NOTE', $file->getPathname(), true);


                // Do khong cÃ³ pháº§n Ä‘áº§u cá»§a Parameters nÃªn nÃ³ bá»? áº£nh hÆ°á»Ÿng bá»Ÿi cÃ¡c Cmd khÃ¡c khÃ´ng pháº£i báº£n thÃ¢n nÃ³ @@ => cáº§n fix
                $fileContent = File::get($file->getPathname());
                for ($idx = 0; $idx < 100; $idx++) {
                    if (preg_match('/Parameters\('.$idx.'\)\.Value = System\.DBNull\.Value/', $fileContent)) {
                        File::replaceInFile('Parameters('.$idx.')', 'Parameters.Add("@p'.$idx.'", SqlDbType.Text)', $file->getPathname());
                    } else {
                        File::replaceInFile('Parameters('.$idx.')', 'Parameters.Add("@p'.$idx.'", SqlDbType.Int)', $file->getPathname());
                    }
                }

                File::replaceInFile('GoSub', 'GoTo', $file->getPathname());

                $arrTBName = ['PRINT', 'PREVIEW', 'CANCEL', 'EXIT', 'EXEC', 'ROWDELETE', 'COPY', 'ROWINSERT', 'EXCEL'];
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

                $this->removeLineByKeySearch('BEGIN TRAN', $file->getPathname(), true, 'dbCon2.BeginTransaction()' . PHP_EOL);
                $this->removeLineByKeySearch('COMMIT TRAN', $file->getPathname(), true, 'dbCon2.Commit()' . PHP_EOL);
                $this->removeLineByKeySearch('ROLLBACK TRAN', $file->getPathname(), true, 'dbCon2.Rollback()' . PHP_EOL);

                $this->removeLineByKeySearch('Dim Index As Short =', $file->getPathname(), true, 'Dim Index As Short = FormUtil.getControlPosition(eventSender)' . PHP_EOL);

                // File::replaceInFile('CellCheck_Numeric(PGrid, ', 'CellCheck_Numeric(', $file->getPathname()); //Sai khi co cac man nhieu grid tren 1 man @@

                File::replaceInFile('VB6.Format(', 'Format(', $file->getPathname());
                File::replaceInFile('VB6.TwipsToPixelsY(', 'TwipsToPixelsY(Me, ', $file->getPathname());
                File::replaceInFile('VB6.TwipsToPixelsX(', 'TwipsToPixelsX(Me, ', $file->getPathname());
                File::replaceInFile('VB6.PixelsToTwipsX(', 'PixelsToTwipsX(Me, ', $file->getPathname());
                File::replaceInFile('VB6.PixelsToTwipsY(', 'PixelsToTwipsY(Me, ', $file->getPathname());
                File::replaceInFile('VB6.TwipsPerPixelX', 'TwipsPerPixelX(Me)', $file->getPathname());
                File::replaceInFile('yyyy/mm/dd', 'yyyy/MM/dd', $file->getPathname());
                File::replaceInFile('hh:nn', 'HH:mm', $file->getPathname());

                File::replaceInFile('CrDraw1', 'mCrDraw', $file->getPathname());

                File::replaceInFile('.RecordCount', '.F ields("rCount").Value', $file->getPathname());

                File::replaceInFile('.set_CellEnabled(', '.set_CellEnable(', $file->getPathname());

                $this->replaceQestionMarkToText($file->getPathname());

                File::replaceInFile('Private mCrForm As CoReports.CrForm', 'Private mCrForm As CrForm' . PHP_EOL . $this->createTab() . 'Private mCrDraw As CrDraw', $file->getPathname());
                File::replaceInFile('If pFncVal <> 0 Then', 'If FormUtil.isPrtEndError(pFncVal) Then', $file->getPathname());

                File::replaceInFile('AxPGRIDLib.AxPerfectGrid', 'CoreLib.UltraGridP', $file->getPathname());
                File::replaceInFile('AxxCBtnLib.AxxCmdBtn', 'CoreLib.ButtonS', $file->getPathname());
                File::replaceInFile('AxxLabelLib.AxxLabel', 'System.Windows.Forms.Label', $file->getPathname());

                $this->replaceInFileWithRegex('pBytes = LenB(StrConv(pPGrid.get_CellText(Row, Col), vbFromUnicode))', 'Dim sutil As StringUtil = New StringUtil(StringUtil.ENC_SHIFTJIS)' . PHP_EOL . $this->createTab(2) . 'pBytes = sutil.getByteCount(pPGrid.get_CellText(Row, Col))' . PHP_EOL, $file->getPathname());

                $this->appendTextToFunction('frm'.$programId.'_Load', $this->createTab(2) .'ImageListUtil.setToolStripImage(Toolbar1)', $file->getPathname(), 'mMBOXTitle = Me.Text');

                File::replaceInFile('ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs', 'ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs', $file->getPathname());
                File::replaceInFile('Handles Me.FormClosed', 'Handles Me.FormClosing', $file->getPathname());

                if (!preg_match('/'.preg_quote('If mnuFILEItem_9.Enabled = False Then', '/').'/', file_get_contents($file->getPathname()))) {
                    $this->appendTextToFunction('frm'.$programId.'_FormClosed', $this->createTab(2) . 'If mnuFILEItem_9.Enabled = False Then' . PHP_EOL . $this->createTab(3) . 'mMsgText = "“o˜^ˆ—’†‚Å‚·BI—¹‚Å‚«‚Ü‚¹‚ñB"' . PHP_EOL . $this->createTab(3) . 'MsgBox(mMsgText, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation,mMBOXTitle)' . PHP_EOL . $this->createTab(3) . 'eventArgs.Cancel = True' . PHP_EOL . $this->createTab(3) . 'Exit Sub' . PHP_EOL . $this->createTab(2) . 'End if' . PHP_EOL, $file->getPathname(), 'Sub frm'.$programId.'_FormClosed');
                }

            }
        }

        return 0;
    }

    protected function createTab($num = 1) {
        return str_repeat(chr(9), $num);
    }

    protected function removeLineByKeySearch($keySearch, $path, $isRegex, $toString = '') {
        $arrFileContent = file($path);
        $exceptText = ['UPGRADE_ISSUE', 'UPGRADE_WARNING', 'UPGRADE_NOTE'];

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

    protected function replaceInFileWithRegex($search, $replace, $path)
    {
        file_put_contents($path, preg_replace('/'.preg_quote($search, '/').'/', $replace, file_get_contents($path), 1));
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
                        $this->replaceInFileWithRegex($arrFileContent[$key - 2] . $arrFileContent[$key - 1] . $arrFileContent[$key], $arrFileContent[$key - 2] . $arrFileContent[$key - 1] . $arrFileContent[$key] . $textAppend . PHP_EOL, $path);
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

    }
}
