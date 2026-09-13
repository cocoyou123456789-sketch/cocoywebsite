import * as T from './vendor/three.module.js';

const host = document.querySelector('.companion-portrait');
const hero = document.querySelector('.hero');
const reduce = matchMedia('(prefers-reduced-motion: reduce)');
try {
  const renderer = new T.WebGLRenderer({ alpha: true, antialias: true });
  renderer.setPixelRatio(Math.min(devicePixelRatio, 2));
  renderer.outputColorSpace = T.SRGBColorSpace;
  renderer.setClearColor(0x141218, 0);
  renderer.domElement.className = 'avatar-canvas';
  const scene = new T.Scene();
  const camera = new T.PerspectiveCamera(32, 1, .1, 50);
  camera.position.set(0, 1.65, 6.5);
  camera.lookAt(0, 1.55, 0);
  scene.add(new T.HemisphereLight(0xffedf1, 0x665676, 2));
  for (const [color, intensity, x, y, z] of [[0xffe4d0, 3, -3, 5, 5], [0xc0bbff, 2, 3, 3, 2], [0xe9a5e2, 3, 1, 4, -3]]) {
    const light = new T.DirectionalLight(color, intensity); light.position.set(x,y,z); scene.add(light);
  }
  const material = (color, roughness=.55) => new T.MeshStandardMaterial({color, roughness});
  const skin=material('#d9a183'), hair=material('#302329', .38), jacket=material('#394651'), lapel=material('#56636d'), shirt=material('#f2e9df'), white=material('#fff7ef'), iris=material('#5d3428', .3), black=material('#181116'), lip=material('#ac5860'), gold=material('#d9b57a', .28);
  const root = new T.Group(); scene.add(root);
  const oval=(parent, mat, x,y,z, sx,sy,sz)=>{
    const m=new T.Mesh(new T.SphereGeometry(1,48,32),mat); m.position.set(x,y,z); m.scale.set(sx,sy,sz); parent.add(m); return m;
  };
  const curve=(parent, points, radius, mat)=>{
    const path=new T.CatmullRomCurve3(points.map(p=>new T.Vector3(...p)));
    const mesh=new T.Mesh(new T.TubeGeometry(path,32,radius,10,false),mat); parent.add(mesh); return mesh;
  };
  // Tailored bust, ivory blouse, and separate lapels.
  oval(root,jacket,0,.15,0,1.03,.95,.44);
  oval(root,shirt,0,.51,.38,.38,.57,.08);
  for(const side of [-1,1]) {
    const shoulder=oval(root,jacket,side*.88,.27,0,.32,.68,.35); shoulder.rotation.z=side*.22;
    const l=oval(root,lapel,side*.37,.55,.42,.14,.61,.045); l.rotation.z=-side*.36;
    oval(root,gold,side*.27,.15,.47,.035,.035,.025);
  }
  oval(root,skin,0,1.09,0,.23,.45,.23);
  // All facial features and hair rotate together around the neck joint.
  const head=new T.Group(); head.position.set(0,1.23,0); root.add(head);
  oval(head,hair,0,.73,-.2,.66,.85,.52);
  oval(head,skin,0,.72,.08,.59,.78,.49);
  const crown=new T.Mesh(new T.SphereGeometry(1,48,24,0,Math.PI*2,0,.88),hair);
  crown.position.set(0,.72,.065);crown.scale.set(.625,.83,.54);head.add(crown);
  for(const side of [-1,1]) {
    oval(head,skin,side*.59,.65,.02,.115,.2,.1);
    oval(head,gold,side*.62,.45,.12,.055,.07,.035);
    oval(head,white,side*.235,.84,.482,.17,.105,.055);
    const eye=new T.Group(); eye.position.set(side*.235,.84,.526); head.add(eye);
    oval(eye,iris,0,0,0,.071,.078,.025);
    oval(eye,black,0,0,.022,.033,.043,.012);
    oval(eye,white,-.019,.026,.033,.016,.017,.008);
    eye.userData.eye=true;
    curve(head,[[side*.075,.99,.46],[side*.22,1.06,.477],[side*.38,1.015,.44]],.022,hair);
    curve(head,[[side*.08,.87,.52],[side*.22,.943,.52],[side*.4,.87,.485]],.012,hair);
    // Long sculpted locks, kept behind the eyes.
    for(let i=0;i<5;i++) {
      const x=side*(.53+i*.042);
      curve(head,[[side*.35,1.36,-.06],[x,1.12,-.05],[x+side*.035,.6,-.02],[x+side*.13,.12,-.05],[x+side*.06,-.38,-.08]],.075,hair);
    }
  }
  oval(head,skin,0,.67,.52,.075,.13,.105);
  curve(head,[[-.18,.43,.511],[-.09,.395,.537],[0,.39,.544],[.09,.395,.537],[.18,.43,.511]],.023,lip);
  // Side-parted swept fringe above the brow.
  for(let i=0;i<7;i++) {
    curve(head,[[.3-i*.07,1.39,.04],[.02-i*.09,1.37,.29],[-.38-i*.023,1.15,.34],[-.55-i*.012,.97,.19]],.075,hair);
  }
  const eyes=[]; head.traverse(o=>{if(o.userData.eye)eyes.push(o);});
  let targetX=0,targetY=0,frame=0,last=0,visible=true;
  const draw=time=>{
    frame=0;
    const dt=Math.min((time-last)/1000||.016,.05); last=time;
    const k=1-Math.exp(-dt*9);
    const yaw=reduce.matches?0:targetX*.65, pitch=reduce.matches?0:targetY*.32;
    head.rotation.y+=(yaw-head.rotation.y)*k;
    head.rotation.x+=(pitch-head.rotation.x)*k;
    head.rotation.z+=(-targetX*.035-head.rotation.z)*k;
    eyes.forEach(eye=>{eye.rotation.y=head.rotation.y*.15; eye.rotation.x=head.rotation.x*.1;});
    renderer.render(scene,camera);
    if(visible && (Math.abs(head.rotation.y-yaw)+Math.abs(head.rotation.x-pitch)>.0001)) frame=requestAnimationFrame(draw);
  };
  const wake=()=>{if(!frame && visible){last=performance.now();frame=requestAnimationFrame(draw);}};
  hero.addEventListener('pointermove',e=>{
    if(e.pointerType==='touch')return;
    const r=host.getBoundingClientRect();
    targetX=T.MathUtils.clamp((e.clientX-r.left-r.width/2)/(innerWidth*.35),-1,1);
    targetY=T.MathUtils.clamp((e.clientY-r.top-r.height*.32)/(innerHeight*.4),-1,1);wake();
  },{passive:true});
  const reset=()=>{targetX=targetY=0;wake();};
  hero.addEventListener('pointerleave',reset);window.addEventListener('blur',reset);reduce.addEventListener('change',reset);
  new ResizeObserver(()=>{const r=host.getBoundingClientRect();renderer.setSize(r.width,r.height);camera.aspect=r.width/r.height;camera.updateProjectionMatrix();wake();}).observe(host);
  new IntersectionObserver(entries=>{visible=entries[0].isIntersecting;if(visible)wake();else if(frame){cancelAnimationFrame(frame);frame=0;}}).observe(hero);
  host.append(renderer.domElement);host.classList.add('has-3d');wake();
} catch(error) { console.warn('3D avatar unavailable; retaining portrait fallback.',error); }
