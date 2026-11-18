'use strict';
module.exports=(req,res)=>{
  res.setHeader('Cache-Control','no-store');
  res.setHeader('Access-Control-Allow-Origin','https://iamkurakov.github.io');
  if(req.method!=='GET')return res.status(405).json({error:'Method not allowed'});
  return res.status(process.env.A19_GITHUB_TOKEN?200:503).json({service:'a19-shared-gateway',ready:!!process.env.A19_GITHUB_TOKEN});
};
