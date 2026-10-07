const items=[
{date:"06 OCT 2026",category:"Actividad",title:"Actividad de MIDES Florida",excerpt:"Información institucional y materiales disponibles para medios de comunicación.",tag:"Florida"},
{date:"05 OCT 2026",category:"Programa",title:"Nuevas acciones territoriales",excerpt:"Contenido preparado para cobertura periodística y difusión institucional.",tag:"Territorio"},
{date:"03 OCT 2026",category:"Comunicado",title:"Información para medios",excerpt:"Comunicado oficial con recursos multimedia asociados.",tag:"MIDES Florida"}
];
const grid=document.getElementById("newsGrid"),search=document.getElementById("search"),category=document.getElementById("category"),count=document.getElementById("resultCount");
function render(){
 const q=search.value.toLowerCase().trim(),c=category.value;
 const data=items.filter(x=>(!q||Object.values(x).join(" ").toLowerCase().includes(q))&&(!c||x.category===c));
 count.textContent=data.length+" resultado"+(data.length===1?"":"s");
 grid.innerHTML=data.map((x,i)=>`<article class="news-card"><div class="news-image"><span>${x.tag}</span></div><div class="news-body"><div class="meta"><span>${x.date}</span><span>${x.category}</span></div><h3>${x.title}</h3><p>${x.excerpt}</p><a href="comunicado.html?id=${i}">Ver comunicado →</a></div></article>`).join("")||'<div class="empty">No encontramos contenidos con esos criterios.</div>';
}
search.addEventListener("input",render);
category.addEventListener("change",render);
render();