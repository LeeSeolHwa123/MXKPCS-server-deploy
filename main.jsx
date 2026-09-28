import React from "react";
import ReactDOM from "react-dom/client";

const rootElement=document.getElementById("root");
const reactRoot=ReactDOM.createRoot(rootElement);
let appReady=false;

function FatalScreen({error}){
  const message=error?.message||String(error||"알 수 없는 오류");
  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",padding:24,boxSizing:"border-box",background:"#f5f6f6",fontFamily:"Pretendard, Arial, sans-serif",color:"#273030"}}>
      <div style={{width:"min(680px, 100%)",padding:"22px 26px",border:"1px solid #d65a5a",borderRadius:12,background:"#fff",boxShadow:"0 10px 28px rgba(26,32,32,.12)"}}>
        <div style={{fontSize:15,fontWeight:700,color:"#a42d2d"}}>XRF 화면 실행 오류</div>
        <div style={{marginTop:9,fontSize:12,lineHeight:1.65,wordBreak:"break-word"}}>{message}</div>
        <div style={{marginTop:10,fontSize:11,color:"#667070"}}>새로고침 후에도 반복되면 이 문구를 그대로 알려주세요.</div>
        <button type="button" onClick={()=>window.location.reload()} style={{marginTop:14,padding:"8px 14px",border:0,borderRadius:7,background:"#273030",color:"#fff",fontWeight:700,cursor:"pointer"}}>다시 불러오기</button>
      </div>
    </div>
  );
}

class AppErrorBoundary extends React.Component{
  constructor(props){
    super(props);
    this.state={error:null};
  }
  static getDerivedStateFromError(error){
    return {error};
  }
  componentDidCatch(error,info){
    console.error("[XRF render error]",error,info);
  }
  render(){
    return this.state.error?<FatalScreen error={this.state.error}/>:this.props.children;
  }
}

window.addEventListener("error",event=>{
  if(appReady) return;
  reactRoot.render(<FatalScreen error={event.error||event.message}/>);
});
window.addEventListener("unhandledrejection",event=>{
  if(appReady) return;
  reactRoot.render(<FatalScreen error={event.reason}/>);
});

import("./app.jsx")
  .then(({default:App})=>{
    reactRoot.render(
      <React.StrictMode>
        <AppErrorBoundary><App /></AppErrorBoundary>
      </React.StrictMode>
    );
    appReady=true;
  })
  .catch(error=>{
    console.error("[XRF module load error]",error);
    reactRoot.render(<FatalScreen error={error}/>);
  });
