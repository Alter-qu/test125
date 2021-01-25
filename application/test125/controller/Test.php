<?php

namespace app\test125\controller;

use think\Controller;
use think\Db;
use think\Image;
use think\Loader;
use think\Request;

class Test extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
        return view();
    }

    /**
     * 显示创建资源表单页.
     *
     * @return \think\Response
     */
    public function check($book)
    {
        $data=\app\test125\model\Test::where('book',$book)->select();
        if ($data){
            return json(['code'=>500,'msg'=>'已存在','res'=>$data]);
        }else{
            return json(['code'=>200,'msg'=>'ok','res'=>null]);
        }

    }

    /**
     * 保存新建的资源
     *
     * @param  \think\Request  $request
     * @return \think\Response
     */
    public function save(Request $request)
    {
        //接收数据
        $param=$request->param();
        // 获取表单上传文件 例如上传了001.jpg
        $file = request()->file('image');
        // 移动到框架应用根目录/public/uploads/ 目录下
        $info = $file->validate(['size'=>1024*1024*3,'ext'=>'jpeg,png'])->move(ROOT_PATH . 'public' . DS . 'uploads');
        if($info){
            // 成功上传后 获取上传信息
            $param['image']=$info->getSaveName();

        }else{
            // 上传失败获取错误信息
           $this->error($file->getError());
        }
       // print_r($param);
        $result = $this->validate(
          $param,
            [
                'book'  => 'require',
                'author'  => 'require',
            ]);
        if(true !== $result){
            // 验证失败 输出错误信息
           $this->error($result);
        }
//        $image = \think\Image::open('./image.png');
//// 按照原图的比例生成一个最大为150*150的缩略图并保存为thumb.png
//        $image->thumb(150, 150)->save('./thumb.png');
//
//        $res=\app\test125\model\Test::create($param,true);
//        return json(['code'=>200,'msg'=>'ok','res'=>$res]);
        return redirect("Test/show");
    }

    /**
     * 显示指定的资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function show()
    {
        $data=\app\test125\model\Test::where('status',1)->order("price","DESC")->paginate(2);
        return view("show",compact('data'));
    }

    /**
     * 显示编辑资源表单页.
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function edit($id)
    {
        $data=\app\test125\model\Test::find($id)->toArray();
        $data['status']="0";
        $res=\app\test125\model\Test::update($data,true);
       //return json(['code'=>200,'msg'=>'ok','res'=>$res]);
        return redirect("Test/show");
    }

    /**
     * 保存更新的资源
     *
     * @param  \think\Request  $request
     * @param  int  $id
     * @return \think\Response
     */
    public function update(Request $request, $id)
    {
        //
    }

    /**
     * 删除指定资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function delete($id)
    {
        //
    }
    public function export()
    {
        //该处应为从数据库中查出数据，这里仅作演示：
        $data = Db::table('test')->select();
        //因不同需求，文件名不同，需要自己构建
        $fileName = '书籍表';
        //调用公共方法
        $this->excelFileExport($data,$fileName);
    }
    function excelFileExport($data = [],$title='')
    {
        //文件名
        $fileName = $title. '('.date("Y-m-d",time()) .'导出）'. ".xls";
        //加载第三方类库
        Loader::import('PHPExcel.Classes.PHPExcel');
        Loader::import('PHPExcel.Classes.PHPExcel.IOFactory.PHPExcel_IOFactory');
        //实例化excel类
        $excelObj = new \PHPExcel();
        //构建列数--根据实际需要构建即可
        $letter = array('A', 'B', 'C' );
        //表头数组--需和列数一致
        $tableheader = array('书名','作者', '价格');
        //填充表头信息
        for ($i = 0; $i < count($tableheader); $i++) {
            $excelObj->getActiveSheet()->setCellValue("$letter[$i]1", "$tableheader[$i]");
        }
        //循环填充数据
        foreach ($data as $k => $v) {
            $num = $k + 1 + 1;
            //设置每一列的内容
            $excelObj->setActiveSheetIndex(0)
                ->setCellValue('A' . $num, $v['book'])
                ->setCellValue('B' . $num, $v['author'])
                ->setCellValue('C' . $num, $v['price']);


            //设置行高
            $excelObj->getActiveSheet()->getRowDimension($k+4)->setRowHeight(30);
    }
        //以下是设置宽度
        $excelObj->getActiveSheet()->getColumnDimension('A')->setWidth(46);
        $excelObj->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $excelObj->getActiveSheet()->getColumnDimension('C')->setWidth(10);
        $excelObj->getActiveSheet()->getColumnDimension('D')->setWidth(20);

        //设置表头行高
        $excelObj->getActiveSheet()->getRowDimension(1)->setRowHeight(28);
        $excelObj->getActiveSheet()->getRowDimension(2)->setRowHeight(28);
        $excelObj->getActiveSheet()->getRowDimension(3)->setRowHeight(28);

        //设置居中
        $excelObj->getActiveSheet()->getStyle('A1:D1'.($k+2))->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        //所有垂直居中
        $excelObj->getActiveSheet()->getStyle('A1:D1'.($k+2))->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);

        //设置字体样式
        $excelObj->getActiveSheet()->getStyle('A1:D1')->getFont()->setName('黑体');
        $excelObj->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(20);
        $excelObj->getActiveSheet()->getStyle('A1:D1')->getFont()->setBold(true);
        $excelObj->getActiveSheet()->getStyle('A1:D1')->getFont()->setBold(true);
        $excelObj->getActiveSheet()->getStyle('A1:D1')->getFont()->setName('宋体');
        $excelObj->getActiveSheet()->getStyle('A1:D1')->getFont()->setSize(16);
        $excelObj->getActiveSheet()->getStyle('A1:D1'.($k+2))->getFont()->setSize(10);

        //设置自动换行
        $excelObj->getActiveSheet()->getStyle('A1:D1'.($k+2))->getAlignment()->setWrapText(true);

        // 重命名表
        $fileName = iconv("utf-8", "gb2312", $fileName);

        // 设置下载打开为第一个表
        $excelObj->setActiveSheetIndex(0);

        //设置header头信息
        header('Content-Type: application/vnd.ms-excel;charset=UTF-8');
        header("Content-Disposition: attachment;filename='{$fileName}'");
        header('Cache-Control: max-age=0');
        $writer = \PHPExcel_IOFactory::createWriter($excelObj, 'Excel5');
        $writer->save('php://output');
        exit();

    }
}
