<?php

namespace App\Http\Controllers;

use Illuminate\Routing\Controller as BaseController;
use Illuminate\Support\Facades\View;
use \PhpOffice\PhpWord\PhpWord;
class HomeController extends Controller
{
    /**
     * Create a new controller instance.
     *
     * @return void
     */
    public function __construct()
    {

    }
    public function index(){
        return "index";
    }

    public function export(){
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        $phpWord->addFontStyle('myOwnStyle', array('color' => 'FF0000'));
        $phpWord->addParagraphStyle('P-Style', array('spaceAfter' => 95));
        $phpWord->addNumberingStyle(
            'multilevel',
            array(
                'type'   => 'multilevel',
                'levels' => array(
                    array('format' => 'decimal', 'text' => '%1.', 'left' => 360, 'hanging' => 360, 'tabPos' => 360),
                    array('format' => 'upperLetter', 'text' => '%2.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720),
                ),
            )
        );
        $predefinedMultilevel = array('listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER_NESTED);

        $header = array('size' => 16, 'bold' => true);
        $section->addImage(
            'https://www.openmaint.org/++plone++tecnoteca.openmainttheme2019/logo/logo.png',
            array(
                'width'         => 200,
                'height'        => 100,
                'marginTop'     => -1,
                'marginLeft'    => -1,
                'wrappingStyle' => 'behind'
            )
        );
        $fontStyleName = 'oneUserDefinedStyle';
        $phpWord->addFontStyle(
            $fontStyleName,
            array('name' => 'Tahoma', 'size' => 15, 'color' => '1B2232', 'bold' => true)
        );
$section->addText(
    'Job Description',
    $fontStyleName,
    array(
        'textAlignment' => 'center',
        'alignment' => 'center'
    )
);
        // 3. colspan (gridSpan) and rowspan (vMerge)

//        $section->addPageBreak();
//        $section->addText(htmlspecialchars('Table with colspan and rowspan'), $header);
        $fontStyleNamesecond = 'secondUserDefinedStyle';
        $phpWord->addFontStyle(
            $fontStyleNamesecond,
            array('name' => 'Tahoma', 'size' => 13, 'color' => 'FFFFFF', 'bold' => false)
        );
        $styleTable = array('borderSize' => 6, 'borderColor' => '999999');
        $cellRowSpan = array('vMerge' => 'restart', 'valign' => 'center', 'bgColor' => '3260a8');
        $cellRowSpan2 = array('vMerge' => 'restart', 'valign' => 'center');
        $cellRowContinue = array('vMerge' => 'continue');
        $cellColSpan = array('gridSpan' => 3, 'valign' => 'center');
        $cellColSpanall = array('gridSpan' => 4, 'valign' => 'center', 'bgColor' => '3260a8');
        $cellColSpanallnotcolor = array('gridSpan' => 4, 'valign' => 'center');
        $cellHCentered = array('align' => 'center');
        $cellVCentered = array('valign' => 'center');
        $cellVCenteredcolored = array('valign' => 'center', 'bgColor' => '3260a8');

        $phpWord->addTableStyle('Colspan Rowspan', $styleTable);
        $table = $section->addTable('Colspan Rowspan');

        $table->addRow();


        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Job Title'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);

        $table->addCell(2000, $cellRowSpan)->addText(htmlspecialchars('Job Code'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(2000, $cellRowSpan2)->addText(htmlspecialchars('F'), null, $cellHCentered);




        $table->addRow();
        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Grade'), $fontStyleNamesecond, $cellHCentered);

        $table->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellRowContinue);

        $table->addRow();
        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Department'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);
        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Section'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);


        $table->addRow();
        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Qualification'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(4000, $cellColSpan)->addText(htmlspecialchars('E'), null, $cellHCentered);

        $table->addRow();
        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Special Certification'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(4000, $cellColSpan)->addText(htmlspecialchars('E'), null, $cellHCentered);

        $table->addRow();
        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Year Of Experiance'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(4000, $cellColSpan)->addText(htmlspecialchars('E'), null, $cellHCentered);


        $table->addRow();
        $table->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Report To'), $fontStyleNamesecond, $cellHCentered);
        $table->addCell(4000, $cellColSpan)->addText(htmlspecialchars('E'), null, $cellHCentered);


        $section->addTextBreak(1);
        $phpWord->addTableStyle('second', $styleTable);
        $table2 = $section->addTable('second');

        $table2->addRow();
        $table2->addCell(8000, $cellColSpanall)->addText(htmlspecialchars('Basic Function'), $fontStyleNamesecond, $cellHCentered);
        $table2->addRow();
//        $table2->addCell(8000, $cellColSpanallnotcolor)->addText(htmlspecialchars('E'), $fontStyleNamesecond, $cellHCentered);
        $table_cell1 =$table2->addCell(8000, $cellColSpanallnotcolor);
        $table_cell1->addListItem(htmlspecialchars('List Item 1'), 0, 'myOwnStyle', $predefinedMultilevel, 'P-Style');
        $table_cell1->addListItem(htmlspecialchars('List Item 2'), 0, 'myOwnStyle', $predefinedMultilevel, 'P-Style');
        $table_cell1->addListItem(htmlspecialchars('List Item 3'), 0, 'myOwnStyle', $predefinedMultilevel, 'P-Style');

        $section->addTextBreak(1);
        $phpWord->addTableStyle('third', $styleTable);
        $table3 = $section->addTable('third');




        $table3->addRow();
        $table3->addCell(8000, $cellColSpanall)->addText(htmlspecialchars('Minimum Requirement'), $fontStyleNamesecond, $cellHCentered);
        $table3->addRow();
//        $table3->addCell(8000, $cellColSpanallnotcolor)->addListItemRun()->addText(htmlspecialchars('List item 1'), array('bold' => true));
        $table_cell =$table3->addCell(8000, $cellColSpanallnotcolor);
        $table_cell->addListItem(htmlspecialchars('List Item 1'), 0, 'myOwnStyle', $predefinedMultilevel, 'P-Style');
        $table_cell->addListItem(htmlspecialchars('List Item 2'), 0, 'myOwnStyle', $predefinedMultilevel, 'P-Style');
        $table_cell->addListItem(htmlspecialchars('List Item 3'), 0, 'myOwnStyle', $predefinedMultilevel, 'P-Style');




        $section->addTextBreak(1);
        $phpWord->addTableStyle('signiture', $styleTable);
        $table4 = $section->addTable('signiture');

        $table4->addRow();
        $table4->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Prepared date'), $fontStyleNamesecond, $cellHCentered);
        $table4->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);
        $table4->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Approved Date'), $fontStyleNamesecond, $cellHCentered);
        $table4->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);
        $table4->addRow();
        $table4->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Reviewed date'), $fontStyleNamesecond, $cellHCentered);
        $table4->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);
        $table4->addCell(2000, $cellVCenteredcolored)->addText(htmlspecialchars('Last Updated Date'), $fontStyleNamesecond, $cellHCentered);
        $table4->addCell(2000, $cellVCentered)->addText(htmlspecialchars('E'), null, $cellHCentered);

// $view_content = View::make('prints', [])->render();
// \PhpOffice\PhpWord\Shared\Html::addHtml($section, $view_content , false, false);
// Adding Text element with font customized using explicitly created font style object...

// Saving the document as OOXML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
try {
    $objWriter->save(storage_path('TestWordFile.docx'));
} catch (Exception $e) {
}
        // header('Content-Type: application/octet-stream');
        // header("Cache-Control: no-cache, must-revalidate");
        // header("Pragma: no-cache");
        // header("Content-Disposition: attachment; filename=TestWordFile");
return response()->download(storage_path('TestWordFile.docx'));

// Saving the document as ODF file...
// $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'ODText');
// $objWriter->save('helloWorld.odt');

// Saving the document as HTML file...
// $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
// $objWriter->save('helloWorld.html');
        // return "echo";
    }
}
