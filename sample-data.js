/* ===========================================================
   DATOS DE DEMO 100% SINTÉTICOS (inventados).
   No corresponden a personas ni clientes reales.
   Sirven para probar la app sin exponer información.
   =========================================================== */
(function(){
  // PRNG determinístico (sin Math.random para que el demo sea estable)
  let seed = 20260519;
  const rnd = ()=>{ seed = (seed*1103515245 + 12345) & 0x7fffffff; return seed/0x7fffffff; };
  const pick = arr => arr[Math.floor(rnd()*arr.length)];
  const pad = (n,l)=> String(n).padStart(l,'0');

  const equipos = [
    { sup:'Equipo Norte · A. Rivera',   agentes:['Lucía Fernández Soto','Pedro Ramos León','Ana Torres Díaz','Mario Quispe Vega','Rosa Flores Nina','José Mendoza Pari','Carla Núñez Ruiz'] },
    { sup:'Equipo Sur · M. Castro',     agentes:['Tamara Hoyos Cruz','Carlos Reyes Meza','Elena Cruz Roca','Juan Pérez Salas','Diego Salas Mejía','Sofía Ramírez Lazo'] },
    { sup:'Equipo Centro · L. Paredes', agentes:['Brenda Olano Gil','Iván Cabrera Mora','Nadia Sosa Vela','Hugo Medina Ríos','Paola Vega Luna','Raúl Campos Ato','Karen Ríos Mío','Luis Tello Vera'] },
    { sup:'Equipo Lima · R. Aguilar',   agentes:['Gina Calle Mora','Brayan Cordano Paz','Yenni Cadillo Mez','Franco Grimaldo Sá','Mía López Soto','Aldo Bravo Alor'] },
  ];
  const ofrec   = ['REGULAR','REGULAR','REGULAR','REGULAR','REGULAR','MULTILINEA X2','MULTILINEA X2','MULTILINEA X3','CROSS-SELLING'];
  const planes  = ['POWER ILIM 69.9','ENTEL CHIP 22.90','POWER 49.9','PLAN MAX 99.9','LIBRE 39.9'];
  const estados = ['VALIDA','VALIDA','VALIDA','VALIDA','VALIDA','VALIDA','VALIDA','VALIDA','OBSERVADA','PENDIENTE'];
  const productos= ['CHIP','CHIP','CHIP','CHIP','CHIP + ACCESORIO','PACK VEP 12'];
  const fechas  = ['2026-05-19','2026-05-20','2026-05-21','2026-05-22','2026-05-23'];

  const rows = [];
  let id = 500000;
  const dni = {};   // DNI fijo por agente (inventado)
  equipos.forEach(eq=> eq.agentes.forEach(a=>{ dni[a]= '4'+pad(Math.floor(rnd()*9000000)+1000000,7); }));

  fechas.forEach(fecha=>{
    equipos.forEach(eq=>{
      eq.agentes.forEach((ag,ai)=>{
        const skill = 0.9 + 0.14*((eq.agentes.length-ai)); // los primeros venden algo más
        for(let h=9; h<=20; h++){
          let dens = (h>=10&&h<=13)||(h>=16&&h<18) ? 1.6 : (h<10?0.9:1.1);
          if(h>=18) dens = 0.8;       // horas extra (después de 18:00)
          const k = Math.floor(rnd()*skill*dens*2.0);
          for(let j=0;j<k;j++){
            const of = pick(ofrec);
            const isCross = of==='CROSS-SELLING';
            const tipoVenta = isCross
              ? pick(['VENTA REGULAR','PORTABILIDAD POSTPAGO','VENTA REGULAR'])
              : pick(['MIGRACION','MIGRACION','MIGRACION','PORTABILIDAD POSTPAGO']);
            const min = Math.floor(rnd()*60);
            rows.push({
              id: String(id++),
              fecha: fecha,
              hora: pad(h,2)+':'+pad(min,2)+':00',
              campana: 'MIGRACION REGULAR',
              tipoOfrec: of,
              tipoVenta: tipoVenta,
              lineaUpsell: '',
              lineaMigrar: tipoVenta==='MIGRACION' ? '9'+pad(Math.floor(rnd()*90000000)+10000000,8) : '',
              plan: pick(planes),
              cargoFijo: (22 + Math.floor(rnd()*78)) + '.90',
              ganancia: (10 + Math.floor(rnd()*40)).toFixed(2),
              cantFamilia: '0',
              tipoProducto: isCross ? pick(productos) : 'CHIP',
              cantAcces: '0',
              montoFinan: '0.00',
              docVendedor: dni[ag],
              supervisor: eq.sup,
              vendedor: ag,
              estado: pick(estados),
              subEstado: ''
            });
          }
        }
      });
    });
  });
  window.SAMPLE_ROWS = rows;
})();
