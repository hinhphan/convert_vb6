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

        // 2-4. [プログラムID].vbproj.userを削除する。
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
        // File::replaceInFile(['[プログラムID]'], $programId, $this->dirVBNET . DIRECTORY_SEPARATOR . $programId . DIRECTORY_SEPARATOR . 'Bas_'.$programId.'.vb');
        
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

                // $this->removeLineByKeySearch('Microsoft\.VisualBasic\.Compatibility\.VB6\.ToolStripMenuItemArray', $file->getPathname(), true);

                // $this->removeLineByKeySearch('System\.Windows\.Forms\.ToolStripSeparator', $file->getPathname(), true);

                $this->removeLineByKeySearch('CrDraw1', $file->getPathname(), true);

                File::replaceInFile('Friend WithEvents lblNMGB As System.Windows.Forms.Label', 'Friend WithEvents lblNMGB As CoreLib.LabelFaculty', $file->getPathname());
                File::replaceInFile('Me.lblNMGB = New System.Windows.Forms.Label', 'Me.lblNMGB = New CoreLib.LabelFaculty', $file->getPathname());
            }
            elseif (preg_match('/^'.$programId.'.*\.vb/', $file->getFilename())) {
                // For logic file
                // $arrFileContent = file($file->getPathname());
                // $tmpIdx = null;

                // foreach ($arrFileContent as $key => $content) {
                //     if (preg_match('/Friend Class frm'.$programId.'/', $content)) {

                //     }
                // }

                File::replaceInFile('System.Windows.Forms.Form', 'Frm_Core', $file->getPathname());

                $this->removeLineByKeySearch('UPGRADE_ISSUE', $file->getPathname(), true);
                $this->removeLineByKeySearch('UPGRADE_WARNING', $file->getPathname(), true);
                $this->removeLineByKeySearch('UPGRADE_NOTE', $file->getPathname(), true);

                $fileContent = File::get($file->getPathname());
                for ($idx = 0; $idx < 30; $idx++) {
                    if (preg_match('/Parameters\('.$idx.'\)\.Value = System\.DBNull\.Value/', $fileContent)) {
                        File::replaceInFile('Parameters('.$idx.')', 'Parameters.Add("@p'.$idx.'", SqlDbType.Text)', $file->getPathname());
                    } else {
                        File::replaceInFile('Parameters('.$idx.')', 'Parameters.Add("@p'.$idx.'", SqlDbType.Int)', $file->getPathname());
                    }
                }

                //File::replaceInFile('GoSub', 'GoTo', $file->getPathname());

                $arrTBName = ['PRINT', 'PREVIEW', 'CANCEL', 'EXIT', 'EXEC', 'ROWDELETE', 'COPY'];
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

                $this->removeLineByKeySearch('BEGIN TRAN', $file->getPathname(), true, 'dbCon2.BeginTransaction()');
                $this->removeLineByKeySearch('COMMIT TRAN', $file->getPathname(), true, 'dbCon2.Commit()');
                $this->removeLineByKeySearch('ROLLBACK TRAN', $file->getPathname(), true, 'dbCon2.Rollback()');

                $this->removeLineByKeySearch('Dim Index As Short =', $file->getPathname(), true, 'Dim Index As Short = FormUtil.getControlPosition(eventSender)');

                File::replaceInFile('CellCheck_Numeric(PGrid, ', 'CellCheck_Numeric(', $file->getPathname());

                File::replaceInFile('VB6.Format(', 'Format(', $file->getPathname());
                File::replaceInFile('yyyy/mm/dd', 'yyyy/MM/dd', $file->getPathname());
                File::replaceInFile('hh:nn', 'HH:mm', $file->getPathname());

                File::replaceInFile('CrDraw1', 'mCrDraw', $file->getPathname());

                File::replaceInFile('.RecordCount', '.F ields("rCount").Value', $file->getPathname());

            }
        }

        return 0;
    }

    protected function removeLineByKeySearch($keySearch, $path, $isRegex, $toString = '') {
        $arrFileContent = file($path);

        foreach ($arrFileContent as $content) {
            if ($isRegex) {
                if (preg_match('/.*'.$keySearch.'.*/', $content)) {
                    File::replaceInFile($content, $toString, $path);
                }
            } else {
                // without regex

            }
        }
    }
}
