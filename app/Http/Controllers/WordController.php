<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\Color;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\TablePosition;
use PhpOffice\PhpWord\Style\Cell;
use PhpOffice\PhpWord\Shared\Html;
use PhpOffice\PhpWord\Style\Line;
use PhpOffice\PhpWord\Shared\Converter;

class WordController extends Controller
{
    public function form_2()
    {
        // สร้างเอกสารใหม่
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();

        // เพิ่ม header
        $header = $section->addHeader();
        $header->addPreserveText('{PAGE}', array('color' => '418AB3'), array('alignment' => Jc::RIGHT));

        // สร้าง style สำหรับหัวข้อ
        $phpWord->addTitleStyle('title', array('name' => 'TH SarabunPSK', 'size' => 20, 'bold' => true), array('alignment' => Jc::CENTER));

        // style อักษร
        $textStyle = array('name' => 'TH SarabunPSK', 'size' => 16);

        // เพิ่มหัวข้อ
        $section->addTitle('แบบฟอร์มที่ 1 ข้อมูลการสมัครรางวัล', 'title');
        $section->addText('____________________', $textStyle, array('alignment' => Jc::CENTER));

        // สร้าง style สำหรับตาราง
        $tableStyle = array('borderSize' => 6, 'borderColor' => 'FFFFFF', 'cellMargin' => 0);
        $firstRowStyle = array('bgColor' => '418AB3', 'color' => 'FFFFFF');
        $phpWord->addTableStyle('CustomTable', $tableStyle, $firstRowStyle);

        // เพิ่มตาราง
        $table = $section->addTable('CustomTable');

        // // เพิ่มแถวและเซลล์ในตาราง
        $table->addRow();
        $table->addCell(9000, array('gridSpan' => 2, 'bgColor' => '418AB3'))->addText('ข้อมูลผลงาน', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#FFFFFF'), array('alignment' => Jc::LEFT));

        $table->addRow();
        $table->addCell(4500)->addText('ประเภทการจัดสมัคร', $textStyle);
        $table->addCell(4500)->addText('ประเภทการอ่านบทความ', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('สัดดวกในการให้บริการ', $textStyle);
        $table->addCell(4500)->addText('เครี่อยงไร้พรมแดน ด้านนโยบายสพติค', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('หน่วยงานต้นสังกัด', $textStyle);
        $table->addCell(4500)->addText('สำนักงานปล็กรธทรวงมหาดไทย', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('หน่วยงานที่รับผิดชอบผลงาน', $textStyle);
        $table->addCell(4500)->addText('สำนักงานจังหวัดขอนแก่น', $textStyle);

        // เพิ่ม section ใหม่สำหรับข้อมูลผู้รับผิดชอบผลงาน
        $table->addRow();
        $table->addCell(9000, array('gridSpan' => 2, 'bgColor' => '418AB3'))->addText('ข้อมูลผู้รับผิดชอบผลงาน', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#FFFFFF'), array('alignment' => Jc::LEFT));
        $table->addRow();
        $table->addCell(4500)->addText('ชื่อ-นามสกุล', $textStyle);
        $table->addCell(4500)->addText('นายสมพงษ์ แคล้วคลาด', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('ตำแหน่ง', $textStyle);
        $table->addCell(4500)->addText('ผู้บัญชาการกองทัพบก', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('สำนัก/กอง', $textStyle);
        $table->addCell(4500)->addText('-', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('เบอร์โทรศัพท์', $textStyle);
        $table->addCell(4500)->addText('02-111-1111 ต่อ 123', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('เบอร์โทรศัพท์มือถือ', $textStyle);
        $table->addCell(4500)->addText('081-111-1111', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('อีเมล', $textStyle);
        $table->addCell(4500)->addText('somphong@gmail.com', $textStyle);

        $table->addRow();
        $table->addCell(9000, array('gridSpan' => 2, 'bgColor' => '418AB3'))->addText('ข้อมูลผู้ประสานงาน', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#FFFFFF'), array('alignment' => Jc::LEFT));

        // ผู้ประสานงาน
        $section->addTextBreak(1);
        $table = $section->addTable('CustomTable');
        $table->addRow();
        $table->addCell(9000, array('gridSpan' => 2, 'bgColor' => 'D7E7F0'))->addText('ผู้ประสานงานคนที่ 1', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#000000'), array('alignment' => Jc::LEFT));
        $table->addRow();
        $table->addCell(4500)->addText('ชื่อ-นามสกุล', $textStyle);
        $table->addCell(4500)->addText('นายสมพงษ์ แคล้วคลาด', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('ตำแหน่ง', $textStyle);
        $table->addCell(4500)->addText('ผู้บัญชาการกองทัพบก', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('สำนัก/กอง', $textStyle);
        $table->addCell(4500)->addText('-', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('เบอร์โทรศัพท์', $textStyle);
        $table->addCell(4500)->addText('02-111-1111 ต่อ 123', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('เบอร์โทรศัพท์มือถือ', $textStyle);
        $table->addCell(4500)->addText('081-111-1111', $textStyle);


        // ผู้ประสานงาน
        $section->addTextBreak(1);
        $table = $section->addTable('CustomTable');
        $table->addRow();
        $table->addCell(9000, array('gridSpan' => 2, 'bgColor' => 'D7E7F0'))->addText('ผู้ประสานงานคนที่ 2', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#000000'), array('alignment' => Jc::LEFT));
        $table->addRow();
        $table->addCell(4500)->addText('ชื่อ-นามสกุล', $textStyle);
        $table->addCell(4500)->addText('นายสมพงษ์ แคล้วคลาด', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('ตำแหน่ง', $textStyle);
        $table->addCell(4500)->addText('ผู้บัญชาการกองทัพบก', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('สำนัก/กอง', $textStyle);
        $table->addCell(4500)->addText('-', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('เบอร์โทรศัพท์', $textStyle);
        $table->addCell(4500)->addText('02-111-1111 ต่อ 123', $textStyle);

        $table->addRow();
        $table->addCell(4500)->addText('เบอร์โทรศัพท์มือถือ', $textStyle);
        $table->addCell(4500)->addText('081-111-1111', $textStyle);

        // เพิ่ม footer
        $imagePath = public_path('pdf\Logo_Footer.jpg');

        $footer = $section->addFooter();
        $footer->addImage($imagePath, array(
            'alignment' => Jc::CENTER,
            'height' => 30
        ));


        // Save the document
        $filename = 'example.docx';
        $phpWord->save(storage_path($filename));

        // Download the document
        return response()->download(storage_path($filename))->deleteFileAfterSend();
    }

    public function org() {
        $phpWord = new PhpWord();

        // สร้าง style สำหรับหัวข้อ
        $phpWord->addTitleStyle('title', array('name' => 'TH SarabunPSK', 'size' => 20, 'bold' => true), array('alignment' => Jc::CENTER));
        // style อักษร
        $textStyle = array('name' => 'TH SarabunPSK', 'size' => 9);

        // ตั้งค่ากระดาษเป็นแนวนอน
        $section = $phpWord->addSection([
            'orientation' => 'landscape'
        ]);

        // Add header with page number
        $header = $section->addHeader();
        $header->addPreserveText('{PAGE}', array('color' => '418AB3'), array('alignment' => Jc::RIGHT));

        // เพิ่ม footer
        $imagePath = public_path('pdf\Logo_Footer.jpg');
        $footer = $section->addFooter();
        $footer->addImage($imagePath, array(
            'alignment' => Jc::CENTER,
            'height' => 30
        ));

        // Add title and subtitle
        $section->addTitle('แบบฟอร์มที่ 2 สรุปลักษณะสำคัญขององค์กร', 'title');
        $section->addText('____________________', $textStyle, ['alignment' => Jc::CENTER]);

        $filler = '<p style="font-family: TH SarabunPSK; font-size: 9pt;">
            ผู้ส่งมอบ พันธมิตร และผู้ให้ความร่วมมือ
            ผู้ส่งมอบ บริษัทผู้ผลิตเครื่องหมายแสดงการเสียภาษีรถ บริษัทผลิตใบอนุญาตขับรถแบบพลาสติก ตรอ. บริษัทจำหน่ายเครื่องมือตรวจสภาพรถ โรงเรียนสอนขับรถที่ได้รับการรับรอง
            พันธมิตร กรมการปกครอง กรมศุลกากร สำนักงานตำรวจแห่งชาติ สำนักงานคณะกรรมการกำกับและส่งเสริมการประกอบธุรกิจประกันภัย หน่วยรับชำระภาษีรถ สถานศึกษาที่ลงนามบันทึกความเข้าใจ สมาคม/มูลนิธิด้านการขนส่งและด้านความปลอดภัยทางถนนต่างๆ
            ผู้ให้ความร่วมมือ ภาคีเครือข่าย “ตรวจรถฟรีขับขี่ปลอดภัย” เช่น ผู้ผลิต/ผู้แทนจำหน่ายรถยนต์ บริษัทประกันภัย ตรอ. เป็นต้น
            ผู้มีส่วนได้ส่วนเสีย
            ประชาชนทั่วไป (ผู้โดยสาร ประชาชน เจ้าของสินค้า (ผู้ใช้บริการขนส่งสินค้า))
            ความต้องการ (ผู้มีส่วนได้ส่วนเสีย) ความปลอดภัยในชีวิตและทรัพย์สิน ความสะดวกและสามารถเข้าถึงบริการได้ง่าย ค่าโดยสาร-ค่าบริการที่เป็นธรรม การได้รับบริการตามมาตรฐานที่กำหนด การบังคับใช้กฎหมายที่เป็นธรรม การตรงต่อเวลาของการให้บริการของรถสาธารณะ ความโปร่งใสในการดำเนินการ การเยียวยาผู้ได้รับผลกระทบจากการใช้รถใช้ถนน
            สมรรถนะหลักขององค์กร พัฒนามาตรฐานเกี่ยวกับการควบคุม กำกับ ดูแลระบบการขนส่งทางถนน (ความปลอดภัยของรถยนต์ คุณภาพของคนขับรถ การขนส่งผู้โดยสาร และการขนส่งสินค้า)
            แหล่งข้อมูลเชิงเปรียบเทียบ องค์การอนามัยโลก (World Health Organization : WHO),ธนาคารโลก (World Bank), IMD World Competitiveness Center,
            Council of Supply Chain Management Professionals (CSCMP’s State of Logistics Report 2019), International Journal of Business and Management Invention
            การเปลี่ยนแปลงความสามารถในการแข่งขัน กรมการขนส่งทางบกมุ่งเน้นพันธกิจด้านการควบคุม กำกับ ดูแลระบบการขนส่งทางถนน และการพัฒนาระบบการขนส่งทางถนนที่สมดุล เพื่อให้การขนส่งทางถนนมีความปลอดภัยและระบบการขนส่งสาธารณะมีคุณภาพ โดยยังคงเน้นจุดแข็งในด้านการบริการที่เป็นเลิศ
            พันธกิจ
            1. พัฒนาระบบควบคุม กำกับ ดูแลระบบการขนส่งทางถนนให้ได้มาตรฐานและมีความปลอดภัย รวมถึงเชื่อมโยงกับการขนส่งรูปแบบอื่น
            2. พัฒนานวัตกรรมการควบคุม กำกับ ดูแล ระบบการขนส่งทางถนนและบังคับใช้กฎหมาย
            3. พัฒนาและส่งเสริมการให้บริการระบบการขนส่งทางถนนให้มีคุณภาพและมีสำนึกรับผิดชอบ
            4. บริหารจัดการองค์กรตามหลักธรรมาภิบาล
            วิสัยทัศน์
            เป็นองค์กรแห่งนวัตกรรมในการควบคุม กำกับ ดูแล ระบบการขนส่งทางถนน ให้มีคุณภาพและปลอดภัย
            ค่านิยม
            ONE DLT (เป้าหมายชัดเจน มีบูรณาการ งานโดดเด่น เน้นเทคโนโลยีดิจิทัล กำกับตามกฎหมาย โปร่งใสเป็นธรรม)
            วัฒนธรรมองค์การ
            มีวัฒนธรรมการทำงานอย่างทุ่มเท เสียสละ สามัคคี มีจิตบริการ (Service Mind)
            งบประมาณ
            ปีงบประมาณ พ.ศ. 2564 ได้รับจัดสรรงบประมาณจำนวน 3,716.07 ล้านบาท
            รายได้
            5,358.81 ล้านบาท
            จำนวนบุคลากร
            5,471 คน (ณ 30 ก.ย. 64) ประกอบด้วย ข้าราชการจำนวน 3,678 คน (ร้อยละ 67.23) อายุ
            </p>
        ';

        // Three columns
        $section = $phpWord->addSection(
            [
                'colsNum' => 3,
                'colsSpace' => 720,
                'breakType' => 'continuous',
                'orientation' => 'landscape'
            ]
        );

        // $section->addText("{$filler}", $textStyle);
        \PhpOffice\PhpWord\Shared\Html::addHtml($section, $filler, false, false);

        // Save the document
        $filename = 'example.docx';
        $phpWord->save(storage_path($filename), 'Word2007');

        // Download the document
        return response()->download(storage_path($filename))->deleteFileAfterSend();
    }

    public function assessment() {
        $phpWord = new PhpWord();

        // สร้าง style สำหรับหัวข้อ
        $phpWord->addTitleStyle('titleHead', array('name' => 'TH SarabunPSK', 'size' => 20, 'bold' => true), array('alignment' => Jc::CENTER));
        $phpWord->addTitleStyle('titleColor', array('name' => 'TH SarabunPSK', 'size' => 16, 'bold' => true, 'bgColor' => 'FFFF00'), array('alignment' => Jc::CENTER));
        $phpWord->addTitleStyle('title', array('name' => 'TH SarabunPSK', 'size' => 16, 'bold' => true), array('alignment' => Jc::LEFT));

        // style อักษร
        $textStyle = array('name' => 'TH SarabunPSK', 'size' => 16);

        // ตั้งค่ากระดาษเป็นแนวนอน
        $section = $phpWord->addSection([
            'orientation' => 'landscape'
        ]);

        // Add header with page number
        $header = $section->addHeader();
        $header->addPreserveText('{PAGE}', array('color' => '418AB3'), array('alignment' => Jc::RIGHT));

        // เพิ่ม footer
        $imagePath = public_path('pdf\Logo_Footer.jpg');
        $footer = $section->addFooter();
        $footer->addImage($imagePath, array(
            'alignment' => Jc::CENTER,
            'height' => 30
        ));

        $fancyTableStyle = ['borderSize' => 6, 'borderColor' => '999999'];
        $cellRowSpan = ['vMerge' => 'restart', 'valign' => 'center'];
        $cellRowContinue = ['vMerge' => 'continue'];
        $cellColSpan = ['gridSpan' => 6, 'valign' => 'center'];
        $cellColSpanTitle = ['gridSpan' => 7, 'valign' => 'center'];
        $cellHCentered = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 0];
        $cellLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT];
        $cellVCentered = ['valign' => 'center'];

        $spanTableStyleName = 'Colspan Rowspan';
        $phpWord->addTableStyle($spanTableStyleName, $fancyTableStyle);

        // Add title and subtitle
        $section->addTitle('แบบฟอร์มที่ 3 แบบประเมินความพร้อมขอรับรางวัล PMQA', 'titleHead');
        $section->addText('____________________', $textStyle, ['alignment' => Jc::CENTER]);

        // สร้าง style สำหรับตาราง
        $tableStyle = array('borderSize' => 6, 'borderColor' => 'FFFFFF', 'cellMargin' => 0);
        $firstRowStyle = array('bgColor' => '418AB3', 'color' => 'FFFFFF');
        $phpWord->addTableStyle('tableHead', $tableStyle, $firstRowStyle);

        $tableHead = $section->addTable('tableHead');
        $tableHead->addRow();
        $tableHead->addCell(14000, array('bgColor' => '418AB3'))
            ->addText('หมวด 1 ด้านการนำองค์การและความรับผิดชอบต่อสังคม', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#FFFFFF'), array('alignment' => Jc::LEFT, 'spaceAfter' => 0));

        $section->addTitle('1.1 การนำองค์การโดยผู้บริหารของส่วนราชการ', 'title');

        $table = $section->addTable($spanTableStyleName);

        // หัวตาราง
        $table->addRow();
        $cell1 = $table->addCell(10000, $cellRowSpan);
        $textrun1 = $cell1->addTextRun($cellHCentered);
        $textrun1->addText('คำถาม', $textStyle);

        $cell2 = $table->addCell(4000, $cellColSpan);
        $textrun2 = $cell2->addTextRun($cellHCentered);
        $textrun2->addText('คะแนน', $textStyle);

        $table->addRow();
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellVCentered)->addText('0', $textStyle, $cellHCentered);
        $table->addCell(null, $cellVCentered)->addText('1', $textStyle, $cellHCentered);
        $table->addCell(null, $cellVCentered)->addText('2', $textStyle, $cellHCentered);
        $table->addCell(null, $cellVCentered)->addText('3', $textStyle, $cellHCentered);
        $table->addCell(null, $cellVCentered)->addText('4', $textStyle, $cellHCentered);
        $table->addCell(null, $cellVCentered)->addText('5', $textStyle, $cellHCentered);
        // หัวตาราง

        $table->addRow();
        $table->addCell(null, $cellColSpanTitle)->addText('ก. วิสัยทัศน์ ค่านิยม', $textStyle, array('align' => 'left', 'spaceAfter' => 0));

        $table->addRow();
        $table->addCell(null, null)->addText('1.วิสัยทัศน์และค่านิยม', $textStyle, array('align' => 'left', 'spaceAfter' => 0));
        $table->addCell(null, ['gridSpan' => 6, 'valign' => 'center'])->addText(null, $textStyle, array('align' => 'left', 'spaceAfter' => 0));

        $table->addRow();
        $table->addCell(null, null)->addText('ผู้บริหารของส่วนราชมีส่วนร่วมในการดําเนินการกําหนดวิสัยทัศน์และค่านิยม', $textStyle, array('align' => 'left', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[x]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));

        $table->addRow();

        // ul
        $texUlLines = [
            'ผู้บริหารของส่วนราชการดําเนินการในเรื่องดังต่อไปนี้',
            '• ความรับผิดชอบต่อการปฏิบัติงานของส่วนราชการ',
            '• ความรับผิดชอบด้านการเงิน และการป้องกันการทุจริตและประพฤติมิชอบ'
        ];

        $cell = $table->addCell(null, null);
        $textRun = $cell->addTextRun(array('align' => 'left', 'spaceAfter' => 0));

        foreach ($texUlLines as $line) {
            $textRun->addText($line, $textStyle);
            $textRun->addTextBreak();
        }
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[ ]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('[x]', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        // ul

        // footer
        $table->addRow();
        $table->addCell(null, null)->addText('Average', array('name' => 'TH SarabunPSK', 'size' => 16, 'bold' => true), array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, ['gridSpan' => 6, 'valign' => 'center'])->addText('4.83', array('name' => 'TH SarabunPSK', 'size' => 16, 'bold' => true), array('align' => 'center', 'spaceAfter' => 0));
        $table->addRow();
        $table->addCell(null, null)->addText('Average Category', array('name' => 'TH SarabunPSK', 'size' => 16, 'bold' => true), array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, ['gridSpan' => 6, 'valign' => 'center'])->addText('4.92', array('name' => 'TH SarabunPSK', 'size' => 16, 'bold' => true), array('align' => 'center', 'spaceAfter' => 0));
        // footer

        $filename = 'example.docx';
        $phpWord->save(storage_path($filename), 'Word2007');

        // Download the document
        return response()->download(storage_path($filename))->deleteFileAfterSend();
    }

    public function executive () {
        // สร้างเอกสารใหม่
        $phpWord = new PhpWord();
        // สร้าง style สำหรับหัวข้อ
        $phpWord->addTitleStyle('title', array('name' => 'TH SarabunPSK', 'size' => 20, 'bold' => true), array('alignment' => Jc::CENTER));
        // style อักษร
        $textStyle = array('name' => 'TH SarabunPSK', 'size' => 16);

        // ตั้งค่ากระดาษเป็นแนวนอน
        $section = $phpWord->addSection();

        // Add header with page number
        $header = $section->addHeader();
        $header->addPreserveText('{PAGE}', array('color' => '418AB3'), array('alignment' => Jc::RIGHT));

        // เพิ่ม footer
        $imagePath = public_path('pdf\Logo_Footer.jpg');
        $footer = $section->addFooter();
        $footer->addImage($imagePath, array(
            'alignment' => Jc::CENTER,
            'height' => 30
        ));

        // Add title and subtitle
        $section->addTitle('แบบฟอร์มที่ 4 บทสรุปผู้บริหาร', 'title');
        $section->addText('____________________', $textStyle, ['alignment' => Jc::CENTER]);

        // Define your HTML content
        $fontName = $textStyle['name'];
        $fontSize = $textStyle['size'];

        $phpWord->setDefaultFontName('TH SarabunPSK');
        $phpWord->setDefaultFontSize(16);

        // สร้าง HTML ที่มีการกำหนดสไตล์ในส่วนหัวและหลายแท็ก <p>
        $html = '
        <html>
        <head>
            <style>
                p {
                    font-family: "' . $fontName . '";
                    font-size: ' . $fontSize . 'pt;
                }
            </style>
        </head>
        <body>
            <p>นี่คือข้อความที่มีฟอนต์และขนาดที่กำหนด</p>
            <p>ข้อความที่สอง</p>
            <p>ข้อความที่สาม</p>
        </body>
        </html>
        ';

        // เพิ่ม HTML ลงในเอกสาร
        Html::addHtml($section, $html, false, false);

        $html .= '
        <p>สำนักงานปลัดกระทรวงอุตสาหกรรม (สปอ.) เป็นส่วนราชการในสังกัดกระทรวงอุตสาหกรรม (อก.) หรือ Ministry of Industry: MIND มีภารกิจตามกฎกระทรวงแบ่งส่วนราชการเกี่ยวกับการพัฒนายุทธศาสตร์</p><figure class="image"><img  style="aspect-ratio:676/481;" src="https://assets.opdc.go.th/opdc-production/1712046584/4weVi4YsrRsZrHQv6ma7RDH7qTt0Z6b2YK6lfcOs.jpg" width="676" height="481"></figure>';

        // $html = $this->addSelfClosingSlash($html);
        // // Add HTML content to the section
        // Html::addHtml($section, $html, false, false);

        // Save the document
        $filename = 'example.docx';
        $phpWord->save(storage_path($filename), 'Word2007');

        // Download the document
        return response()->download(storage_path($filename))->deleteFileAfterSend();
    }

    // แบบฟอร์มที่ 5
    public function appPath () {
        // สร้างเอกสารใหม่
        $phpWord = new PhpWord();
        // สร้าง style สำหรับหัวข้อ
        $phpWord->addTitleStyle('title', array('name' => 'TH SarabunPSK', 'size' => 20, 'bold' => true), array('alignment' => Jc::CENTER));
        // style อักษร
        $textStyle = array('name' => 'TH SarabunPSK', 'size' => 16);

        // ตั้งค่ากระดาษเป็นแนวนอน
        $section = $phpWord->addSection();

        // Add header with page number
        $header = $section->addHeader();
        $header->addPreserveText('{PAGE}', array('color' => '418AB3'), array('alignment' => Jc::RIGHT));

        // เพิ่ม footer
        $imagePath = public_path('pdf\Logo_Footer.jpg');
        $footer = $section->addFooter();
        $footer->addImage($imagePath, array(
            'alignment' => Jc::CENTER,
            'height' => 30
        ));

        // Add title and subtitle
        $section->addTitle('แบบฟอร์มที่ 5 รายงานผลการดำเนินการพัฒนาองค์การ', 'title');
        $section->addText('____________________', $textStyle, ['alignment' => Jc::CENTER]);

        // สร้าง style สำหรับตาราง
        $tableStyle = array('borderSize' => 6, 'borderColor' => 'FFFFFF', 'cellMargin' => 0);
        $firstRowStyle = array('bgColor' => '418AB3', 'color' => 'FFFFFF');
        $phpWord->addTableStyle('CustomTable', $tableStyle, $firstRowStyle);

        // เพิ่มตาราง
        $table = $section->addTable('CustomTable');

        $table->addRow();
        $table->addCell(9000, array('gridSpan' => 2, 'bgColor' => '418AB3'))->addText('ส่วนที่ 1 ลักษณะสําคัญขององค์การ', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#FFFFFF'), array('alignment' => Jc::LEFT));

        $section->addTextBreak(1);
        $table = $section->addTable('CustomTable');
        $table->addRow();
        $table->addCell(9000, array('gridSpan' => 2, 'bgColor' => 'D7E7F0'))->addText('1.1 ลักษณะองค์การ : คุณลักษณะสําคัญของส่วนราชการคืออะไร', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#000000'), array('alignment' => Jc::LEFT));

        // Add title style
        $groupStyle = array(
            'size' => 16,
            'color' => '418AB3',
            'name' => 'TH SarabunPSK'
        );

        $paragraphStyle = array(
            'borderBottomSize' => 1,
            'borderBottomColor' => '418AB3',
            'spaceBefore' => 10,
            'name' => 'TH SarabunPSK'
        );

        $section->addText('', null, $paragraphStyle);
        $section->addText('ก. สภาพแวดล้อมของส่วนราชการ', $groupStyle, array('alignment' => 'left'));

        // Save the document
        $filename = 'example.docx';
        $phpWord->save(storage_path($filename));

        // Download the document
        return response()->download(storage_path($filename))->deleteFileAfterSend();
    }

    // แบบฟอร์มที่ 5 ตาราง
    public function appPathTable () {
        $phpWord = new PhpWord();

        $section = $phpWord->addSection([
            'orientation' => 'landscape'
        ]);

        // style อักษร
        $textStyle = array('name' => 'TH SarabunPSK', 'size' => 16);

        // สร้าง style สำหรับตาราง
        $tableStyle = array('borderSize' => 6, 'borderColor' => 'FFFFFF', 'cellMargin' => 0);
        $firstRowStyle = array('bgColor' => '418AB3', 'color' => 'FFFFFF');
        $phpWord->addTableStyle('CustomTable', $tableStyle, $firstRowStyle);

        // เพิ่มตาราง
        $table = $section->addTable('CustomTable');

        $table->addRow();
        $table->addCell(14000, array('gridSpan' => 4, 'bgColor' => '418AB3'))->addText('ส่วนที่ 1 ลักษณะสําคัญขององค์การ', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#FFFFFF'), array('alignment' => Jc::LEFT, 'spaceAfter' => 0));

        $section->addTextBreak(1);
        $table = $section->addTable('CustomTable');
        $table->addRow();
        $table->addCell(14000, array('gridSpan' => 2, 'bgColor' => 'D7E7F0'))->addText('1.1 ลักษณะองค์การ : คุณลักษณะสําคัญของส่วนราชการคืออะไร', array('bold' => true, 'name' => 'TH SarabunPSK', 'size' => 16, 'color' => '#000000'), array('alignment' => Jc::LEFT, 'spaceAfter' => 0));

        $phpWord->addTitleStyle('titleHead', array('name' => 'TH SarabunPSK', 'size' => 20, 'bold' => true), array('alignment' => Jc::LEFT));
        $section->addTitle('กลุ่มตัวชี้วัดที่ 1 ตัวชี้วัดด้านผลผลิตและการบริการตามพันธกิจหลักของส่วนราชการ', 'titleHead');

        $fancyTableStyle = ['borderSize' => 6, 'borderColor' => '999999'];
        $cellRowSpan = ['vMerge' => 'restart', 'valign' => 'center', 'exact' => true];
        $cellRowContinue = ['vMerge' => 'continue'];
        $cellColSpan = ['gridSpan' => 3, 'valign' => 'center', 'exact' => true];
        $cellColSpanTitle = ['gridSpan' => 7, 'valign' => 'center'];
        $cellHCentered = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 0];
        $cellLeft = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT];
        $cellVCentered = ['valign' => 'center'];

        $spanTableStyleName = 'Colspan Rowspan';
        $phpWord->addTableStyle($spanTableStyleName, $fancyTableStyle);

        // สร้าง style สำหรับตาราง
        $tableStyle = array('borderSize' => 6, 'borderColor' => 'FFFFFF', 'cellMargin' => 0);
        $firstRowStyle = array('bgColor' => '418AB3', 'color' => 'FFFFFF');
        $phpWord->addTableStyle('tableHead', $tableStyle, $firstRowStyle);

        $table = $section->addTable($spanTableStyleName);

        // หัวตาราง
        $table->addRow();
        $cell1 = $table->addCell(1000, $cellRowSpan);
        $textrun1 = $cell1->addTextRun($cellHCentered);
        $textrun1->addText('ลำดับ', $textStyle);

        $cell2 = $table->addCell(1000, $cellRowSpan);
        $textrun2 = $cell2->addTextRun($cellHCentered);
        $textrun2->addText('ชื่อตัวชี้วัด', $textStyle);

        $cell3 = $table->addCell(1000, $cellRowSpan);
        $textrun3 = $cell3->addTextRun($cellHCentered);
        $textrun3->addText('ค่าน้อยดี/ค่ามากดี', $textStyle);

        $cell4 = $table->addCell(1000, $cellRowSpan);
        $textrun4 = $cell4->addTextRun($cellHCentered);
        $textrun4->addText('ค่าเป้าหมายปีล่าสุด', $textStyle);

        $cell5 = $table->addCell(1000, $cellRowSpan);
        $textrun5 = $cell5->addTextRun($cellHCentered);
        $textrun5->addText('หน่วย', $textStyle);

        $cell6 = $table->addCell(3000, $cellColSpan);
        $textrun6 = $cell6->addTextRun($cellHCentered);
        $textrun6->addText('ข้อมูลย้อนหลังอย่างน้อย 3 จุด', $textStyle);

        $cell7 = $table->addCell(1000, $cellRowSpan);
        $textrun7 = $cell7->addTextRun($cellHCentered);
        $textrun7->addText('หมายเหตุ', $textStyle);

        $cell8 = $table->addCell(1000, $cellRowSpan);
        $textrun8 = $cell8->addTextRun($cellHCentered);
        $textrun8->addText('% ความสำเร็จ', $textStyle);

        $cell9 = $table->addCell(1000, $cellRowSpan);
        $textrun9 = $cell9->addTextRun($cellHCentered);
        $textrun9->addText('คะแนน', $textStyle);

        $table->addRow();
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellVCentered)->addText('2564', $textStyle, $cellHCentered);
        $table->addCell(null, $cellVCentered)->addText('2565', $textStyle, $cellHCentered);
        $table->addCell(null, $cellVCentered)->addText('2566', $textStyle, $cellHCentered);
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellRowContinue);
        $table->addCell(null, $cellRowContinue);
        // หัวตาราง

        $table->addRow();
        $table->addCell(null, null)->addText('1', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('จำนวนข้อเสนอแนะเชิงนโยบายที่สอดคล้องกับปัญหาสาธารณสุขระดับประเทศ', $textStyle, array('align' => 'left', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('ค่ามากดี', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('15', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('เรื่อง', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('15', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('15', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('17', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('เป็นตัวชี้วัดที่ชี้แจงเกี่ยวกับจำนวนข้อ', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('300', $textStyle, array('align' => 'center', 'spaceAfter' => 0));
        $table->addCell(null, null)->addText('100', $textStyle, array('align' => 'center', 'spaceAfter' => 0));

        $filename = 'example.docx';
        $phpWord->save(storage_path($filename), 'Word2007');

        // Download the document
        return response()->download(storage_path($filename))->deleteFileAfterSend();
    }

    public function addSelfClosingSlash($text) {
        // ใช้ regular expression เพื่อหาแท็ก img ที่ไม่มี self-closing slash
        $pattern = '/<img(.*?)>/';
        $replacement = '<img$1 />';

        return preg_replace($pattern, $replacement, $text);
    }
}
