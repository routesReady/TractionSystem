const mongoose=require('mongoose');

const sourceSchema=new mongoose.Schema({
  key:{type:String,unique:true,required:true},
  fileName:String,
  contentType:String,
  fileData:{type:Buffer},
  uploadedAt:{type:Date,default:Date.now},
  rows:{type:[mongoose.Schema.Types.Mixed],default:[]}
},{timestamps:true});

const savedSchema=new mongoose.Schema({
  searchDate:{type:String,required:true},
  weekday:{type:String,required:true},
  sourceRow:{type:Number,required:true},
  addedAt:{type:Date,default:Date.now},
  row:{type:mongoose.Schema.Types.Mixed,required:true}
},{timestamps:true});

savedSchema.index({searchDate:1,sourceRow:1},{unique:true});

module.exports={
  SourceDataset:mongoose.model('SourceDataset',sourceSchema),
  SavedRow:mongoose.model('SavedRow',savedSchema)
};
