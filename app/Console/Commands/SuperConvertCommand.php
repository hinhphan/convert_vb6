<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\File;

class SuperConvertCommand extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'super:convert';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Lets do it...';

    protected const CONVERT_TYPE = [
        'INIT_CONVERT',
        'UPGRADE_VB',
        'UPDATE_DOT_NET',
        'COPY_FILE',
        'COPY_DIR',
        'DELETE_FILE',
        'DELETE_DIR',
        'REPLACE_TEXT',
        'REMOVE_LINE_TEXT',
        'REPLACE_LINE_TO_TEXT',
    ];

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        $this->removeLineByKeySearch('UPGRADE_ISSUE', public_path('vb_test.vb'), true);
        $this->removeLineByKeySearch('UPGRADE_WARNING', public_path('vb_test.vb'), true);
        $this->removeLineByKeySearch('UPGRADE_NOTE', public_path('vb_test.vb'), true);

        $arrFileContent = file(public_path('vb_test.vb'));

        $start = 0;
        $arrCheckToResetStart = ['delete', 'select', 'insert', 'update'];
        $countQuestionMark = substr_count(file_get_contents(public_path('vb_test.vb')), '?');

        foreach ($arrFileContent as $fileContent) {
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
                $this->replaceInFileWithRegex($fileContent, $newFileContent, public_path('vb_test.vb'));
            }

            if ($countQuestionMark <= 0) {
                break;
            }
        }

        return 0;
    }

    protected function replaceInFileWithRegex($search, $replace, $path)
    {
        file_put_contents($path, preg_replace('/'.preg_quote($search, '/').'/', $replace, file_get_contents($path), 1));
    }

    protected function removeLineByKeySearch($keySearch, $path, $isRegex, $toString = '') {
        $arrFileContent = file($path);

        foreach ($arrFileContent as $content) {
            if ($isRegex) {
                if (preg_match('/.*'.$keySearch.'.*/', $content)) { //cho nay nen bo regex 2 ben mac dinh di
                    File::replaceInFile($content, $toString, $path);
                }
            } else {
                // without regex

            }
        }
    }
}
