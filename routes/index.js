const router = require("koa-router")();
const multer = require("koa-multer");
const fs = require("fs");
const path = require("path");
const send = require("koa-send");
const { merge } = require("../excel/merge");
const { export_erp } = require('../excel/erp');
const { years_data } = require('../excel/count')
const { cloud } = require('../excel/cloud')

const storage = multer.diskStorage({
  //文件保存路径
  destination: function (req, file, cb) {
    cb(null, "../upload/");
  },
  //修改文件名称
  filename: function (req, file, cb) {
    var fileFormat = file.originalname.split("."); //以点分割成数组，数组的最后一项就是后缀名
    cb(null, Date.now() + "." + fileFormat[fileFormat.length - 1]);
  },
});
const upload = multer({ storage });
router.post("/upload", async (ctx) => {
  const file = ctx.request.files.file; // 获取上传文件
  const reader = fs.createReadStream(file.path); // 创建可读流 此处需引入 fs模块
  const arr = file.name.split("."); // 获取上传文件扩展名
  const upStream = fs.createWriteStream(
    path.resolve(__dirname, `../upload/${arr[0]}.${arr[1]}`)
  ); // 创建可写流
  reader.pipe(upStream);
  ctx.body = {
    code: 200,
    message: "上传成功",
    success: true,
  };
});

router.post("/turn", async (ctx) => {
  try{
      const {type,filename} = ctx.request.body
      global.currentFileName = filename
      switch(type){
        case 1:
          await merge();
          break;
        case 2:
          await export_erp();
          break;
        case 3:
          await years_data();
          break;
        case 4:
          await cloud();
          break; 
        default:
          break;
      }
      
      ctx.body = {
        code :200,
        message:'转化成功',
        fileName:"对账表.xlsx",
        success:true
      }
  }catch(e){
    throw new Error('转化文件失败')
  }
});

router.get('/download/:name',async ctx => {
  const name = decodeURI(ctx.params.name);
  ctx.attachment(name);
  await send(ctx,name,{root:path.join(__dirname,'../files')})
})

module.exports = router;
