<?php

namespace app\controllers;

use Yii;
use Mpdf\Mpdf;
use yii\web\Controller;
use backend\models\Client;
use yii\base\DynamicModel;
use alarm3d\export\OpenTBS;
use yii\web\NotFoundHttpException;
use backend\models\search\ClientSearch;
/**
 * WordController implements the CRUD actions for  model.
 */
class WordController extends Controller
{
    /**
     * Lists all Client models.
     * @return mixed
     */
    public function actionIndex()
    {
        $searchModel = new ClientSearch();
        $dataProvider = $searchModel->search(Yii::$app->request->queryParams);

        $field = [
            'fileImport' => 'File Import',
        ];
        $modelImport = DynamicModel::validateData($field, [
            [['fileImport'], 'required'],
            [['fileImport'], 'file', 'extensions'=>'xls,xlsx','maxSize'=>1024*1024],
        ]);

        return $this->render('index', [
            'searchModel' => $searchModel,
            'dataProvider' => $dataProvider,
            'modelImport' => $modelImport,
        ]);
    }

    /**
     * Finds the Mahasiswa model based on its primary key value.
     * If the model is not found, a 404 HTTP exception will be thrown.
     * @param integer $id
     * @return Mahasiswa the loaded model
     * @throws NotFoundHttpException if the model cannot be found
     */
    protected function findModel($id)
    {
        if (($model = Client::findOne($id)) !== null) {
            return $model;
        } else {
            throw new NotFoundHttpException('The requested page does not exist.');
        }
    }

    /*
    IMPORT WITH PHPEXCEL
    */
    public function actionImport()
    {
        $field = [
            'fileImport' => 'File Import',
        ];

        $modelImport = DynamicModel::validateData($field, [
            [['fileImport'], 'required'],
            [['fileImport'], 'file', 'extensions'=>'xls,xlsx','maxSize'=>1024*1024],
        ]);

        if (Yii::$app->request->post()) {
            $modelImport->fileImport = \yii\web\UploadedFile::getInstance($modelImport, 'fileImport');
            if ($modelImport->fileImport && $modelImport->validate()) {
                $inputFileType = \PHPExcel_IOFactory::identify($modelImport->fileImport->tempName );
                $objReader = \PHPExcel_IOFactory::createReader($inputFileType);
                $objPHPExcel = $objReader->load($modelImport->fileImport->tempName);
                $sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
                $baseRow = 2;
                while(!empty($sheetData[$baseRow]['A'])){
                    $model = new Client();
                    $model->name = (string)$sheetData[$baseRow]['B'];
                    $model->code = (string)$sheetData[$baseRow]['C'];
                    $model->save();
                    //die(print_r($model->errors));
                    $baseRow++;
                }
                Yii::$app->getSession()->setFlash('success', 'Success');
            }
            else{
                Yii::$app->getSession()->setFlash('error', 'Error');
            }
        }

        return $this->redirect(['index']);
    }

    /*
	EXPORT WITH PHPEXCEL
	*/
    public function actionExportExcel()
    {
        $searchModel = new ClientSearch();
        $dataProvider = $searchModel->search(Yii::$app->request->queryParams);

        $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
        $template = Yii::getAlias('@vendor/alarm3d/yii2-export').'/templates/phpexcel/export.xlsx';
        $objPHPExcel = $objReader->load($template);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(\PHPExcel_Worksheet_PageSetup::PAPERSIZE_FOLIO);
        $baseRow=2; // line 2
        foreach($dataProvider->getModels() as $mahasiswa){
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$baseRow, $baseRow-1);
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$baseRow, $mahasiswa->name);
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$baseRow, $mahasiswa->code);
            $baseRow++;
        }
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="export.xlsx"');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
        $objWriter->save('php://output');
        exit;
    }

    /*
    EXPORT WITH OPENTBS
    */
    public function actionExportExcel2()
    {
        $searchModel = new ClientSearch();
        $dataProvider = $searchModel->search(Yii::$app->request->queryParams);

        // Initalize the TBS instance
        $OpenTBS = new OpenTBS(); // new instance of TBS
        // Change with Your template kaka
        $template = Yii::getAlias('@vendor/alarm3d/yii2-export').'/templates/opentbs/ms-excel.xlsx';
        $OpenTBS->LoadTemplate($template, OPENTBS_ALREADY_UTF8); // Also merge some [onload] automatic fields (depends of the type of document).
        //$OpenTBS->VarRef['modelName']= "Mahasiswa";
        $data = [];
        $no=1;
        foreach($dataProvider->getModels() as $mahasiswa){
            $data[] = [
                'no'=>$no++,
                'name'=>$mahasiswa->name,
                'code'=>$mahasiswa->code,
            ];
        }

        $data2[0] = [
            'no'=>'X',
            'name'=>'Y',
            'code'=>'Z',
        ];
        $data2[1] = [
            'no'=>'X',
            'name'=>'Y',
            'code'=>'Z',
        ];
        $OpenTBS->MergeBlock('data', $data);
        $OpenTBS->MergeBlock('data2', $data2);
        // Output the result as a file on the server. You can change output file
        $OpenTBS->Show(OPENTBS_DOWNLOAD, 'export.xlsx'); // Also merges all [onshow] automatic fields.
        exit;
    }

    /*
    EXPORT WITH OPENTBS
    */
    public function actionExportWord()
    {
        $searchModel = new ClientSearch();
        $dataProvider = $searchModel->search(Yii::$app->request->queryParams);

        // Initalize the TBS instance
        $OpenTBS = new OpenTBS; // new instance of TBS
        // Change with Your template kaka
        $template = Yii::getAlias('@vendor/alarm3d/yii2-export').'/templates/opentbs/ms-word.docx';
        $OpenTBS->LoadTemplate($template, OPENTBS_ALREADY_UTF8); // Also merge some [onload] automatic fields (depends of the type of document).
        //$OpenTBS->VarRef['modelName']= "Client";
        $data = [];
        $no=1;
        foreach($dataProvider->getModels() as $mahasiswa){
            $data[] = [
                'no'=>$no++,
                'name'=>$mahasiswa->name,
                'code'=>$mahasiswa->code,
            ];
        }
        $OpenTBS->MergeBlock('data', $data);
        // Output the result as a file on the server. You can change output file
        $OpenTBS->Show(OPENTBS_DOWNLOAD, 'export.docx'); // Also merges all [onshow] automatic fields.
        exit;
    }

    /*
    EXPORT WITH MPDF
    */
    public function actionExportPdf()
    {
        $searchModel = new ClientSearch();
        $dataProvider = $searchModel->search(Yii::$app->request->queryParams);//Yii::$app->request->queryParams
        $html = $this->renderPartial('_pdf',['dataProvider'=>$dataProvider]);
        $mpdf=new Mpdf(['c','A4','','' , 0 , 0 , 0 , 0 , 0 , 0]);
        $mpdf->SetDisplayMode('fullpage');
        $mpdf->list_indent_first_level = 0;  // 1 or 0 - whether to indent the first level of a list
        $mpdf->WriteHTML($html);
        $mpdf->Output();
        exit;
    }

}