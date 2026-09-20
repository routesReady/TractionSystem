const mongoose=require('mongoose');

const sourceSchema=new mongoose.Schema({
  key:{type:String,unique:true,required:true}, fileName:String, contentType:String, fileData:{type:Buffer}, uploadedAt:{type:Date,default:Date.now},
  rows:{type:[mongoose.Schema.Types.Mixed],default:[]},
  dropdowns:{worked:{type:[String],default:[]},noReason:{type:[String],default:[]}}
},{timestamps:true});

const savedSchema=new mongoose.Schema({
  searchDate:{type:String,required:true}, weekday:{type:String,required:true}, sourceRow:{type:Number,required:true}, addedAt:{type:Date,default:Date.now},
  row:{type:mongoose.Schema.Types.Mixed,required:true},
  shedRemark:{type:String,default:''},
  oem:{type:String,default:''},
  shedRemarkClosed:{type:Boolean,default:false},
  oemRemarkClosed:{type:Boolean,default:false}
},{timestamps:true});
savedSchema.index({searchDate:1,sourceRow:1},{unique:true});

const userSchema=new mongoose.Schema({
  username:{type:String,required:true,trim:true},
  role:{type:String,enum:['admin','shed','oem'],required:true},
  passwordHash:{type:String,required:true},
  mobileNumber:{type:String,default:''},
  isActive:{type:Boolean,default:true},
  authVersion:{type:Number,default:0},
  mustChangePassword:{type:Boolean,default:false}
},{timestamps:true});
userSchema.index({username:1,role:1},{unique:true});
userSchema.index({username:1});

const passwordResetRequestSchema=new mongoose.Schema({
  requestId:{type:String,unique:true,required:true},
  username:{type:String,required:true},
  role:{type:String,enum:['shed','oem'],required:true},
  mobileNumber:{type:String,required:true},
  status:{type:String,enum:['Pending','Approved','Rejected'],default:'Pending'},
  requestedAt:{type:Date,default:Date.now},
  approvedAt:{type:Date},
  approvedBy:{type:String,default:''},
  smsStatus:{type:String,default:'not_sent'},
  smsError:{type:String,default:''}
},{timestamps:true});
passwordResetRequestSchema.index({status:1,requestedAt:-1});

module.exports={
  SourceDataset:mongoose.model('SourceDataset',sourceSchema),
  SavedRow:mongoose.model('SavedRow',savedSchema),
  User:mongoose.model('User',userSchema),
  PasswordResetRequest:mongoose.model('PasswordResetRequest',passwordResetRequestSchema)
};
