document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const processBtn = document.getElementById('processBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const downloadPdfBtn = document.getElementById('downloadPdfBtn');
    const loaderOverlay = document.getElementById('loaderOverlay');
    const metricOps = document.getElementById('metricOps');
    const metricMachines = document.getElementById('metricMachines');
    const ganttChartDiv = document.getElementById('ganttChart');
    
    // Inject Modal HTML
    const modalHtml = `
        <div id="taskModal" class="modal-overlay">
            <div class="modal-content">
                <span class="modal-close">&times;</span>
                <h3 id="modalTitle">Detalles de Tarea</h3>
                <div class="modal-body" id="modalBody"></div>
            </div>
        </div>
    `;
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    const modal = document.getElementById('taskModal');
    const modalClose = document.querySelector('.modal-close');
    modalClose.onclick = () => modal.classList.remove('active');
    window.onclick = (e) => { if (e.target == modal) modal.classList.remove('active'); }

    // Handle file selection
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            fileName.textContent = e.target.files[0].name;
            processBtn.disabled = false;
        } else {
            fileName.textContent = 'Ningún archivo seleccionado';
            processBtn.disabled = true;
        }
    });

    // State for export
    let globalOpsData = null;
    let globalColorMap = null;

    // Handle processing
    processBtn.addEventListener('click', async () => {
        const file = fileInput.files[0];
        if (!file) return;

        loaderOverlay.classList.add('active');

        const formData = new FormData();
        formData.append('file', file);

        try {
            const response = await fetch('/api/schedule', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error('Error en el procesamiento de datos');
            }

            const data = await response.json();
            
            globalOpsData = data.tasks;
            globalColorMap = data.color_map;
            
            metricOps.textContent = data.stats ? data.stats.total_ops : '0';
            metricMachines.textContent = data.stats ? data.stats.machines_used : '0';
            
            // Enable exports
            downloadBtn.disabled = false;
            downloadPdfBtn.disabled = false;
            
            // Renderizamos la tabla
            renderPlaneadorTable(data.tasks, data.color_map);

        } catch (error) {
            console.error('Error:', error);
            alert('Hubo un error procesando el archivo: ' + error.message);
        } finally {
            loaderOverlay.classList.remove('active');
        }
    });

    // Handle Excel Export
    downloadBtn.addEventListener('click', async () => {
        if(!globalOpsData) return;
        loaderOverlay.classList.add('active');
        document.querySelector('.loader-overlay p').textContent = 'Generando archivo Excel...';
        
        try {
            const res = await fetch('/api/export_excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    tasks: globalOpsData,
                    color_map: globalColorMap
                })
            });
            if (!res.ok) throw new Error('Error del servidor al exportar Excel');
            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Planificacion_ProPlan.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        } catch(e) {
            alert('Error exportando: ' + e.message);
        } finally {
            loaderOverlay.classList.remove('active');
            document.querySelector('.loader-overlay p').textContent = 'Procesando algoritmo de optimización...';
        }
    });

    // Handle PDF Export
    downloadPdfBtn.addEventListener('click', () => {
        const element = document.getElementById('tableWrapper');
        if(!element) return;
        
        loaderOverlay.classList.add('active');
        document.querySelector('.loader-overlay p').textContent = 'Generando PDF (Puede tardar unos segundos)...';
        
        const oldOverflow = element.style.overflow;
        const oldPosition = element.style.position;
        const oldWidth = element.style.width;
        
        element.style.overflow = 'visible'; 
        element.style.position = 'relative';
        element.style.width = 'max-content';
        
        const frozenEls = element.querySelectorAll('.frozen-col');
        frozenEls.forEach(el => el.style.position = 'static');
        
        const opt = {
            margin:       0.2,
            filename:     'Planificacion_ProPlan.pdf',
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2, scrollX: 0, scrollY: 0, windowWidth: element.scrollWidth, width: element.scrollWidth },
            jsPDF:        { unit: 'in', format: 'a3', orientation: 'landscape' }
        };

        html2pdf().set(opt).from(element).save().then(() => {
            element.style.overflow = oldOverflow;
            element.style.position = oldPosition;
            element.style.width = oldWidth;
            frozenEls.forEach(el => el.style.position = 'sticky');
            
            loaderOverlay.classList.remove('active');
            document.querySelector('.loader-overlay p').textContent = 'Procesando algoritmo de optimización...';
        }).catch(err => {
            console.error(err);
            alert("Error al generar PDF: " + err);
            
            element.style.overflow = oldOverflow;
            element.style.position = oldPosition;
            element.style.width = oldWidth;
            frozenEls.forEach(el => el.style.position = 'sticky');
            loaderOverlay.classList.remove('active');
        });
    });

    function getContrastYIQ(hexcolor){
        if (!hexcolor) return '#ffffff';
        hexcolor = hexcolor.replace("#", "");
        var r = parseInt(hexcolor.substr(0,2),16);
        var g = parseInt(hexcolor.substr(2,2),16);
        var b = parseInt(hexcolor.substr(4,2),16);
        var yiq = ((r*299)+(g*587)+(b*114))/1000;
        return (yiq >= 128) ? '#0f172a' : '#ffffff';
    }

    // Export function to window so inline onclick works
    window.showTaskDetails = function(op, cliente, desc, machine, start, end) {
        document.getElementById('modalTitle').textContent = `OP: ${op}`;
        document.getElementById('modalBody').innerHTML = `
            <p><strong>Cliente:</strong> ${cliente}</p>
            <p><strong>Descripción:</strong> ${desc}</p>
            <p><strong>Máquina Asignada:</strong> ${machine}</p>
            <p><strong>Fecha Inicio:</strong> ${start}</p>
            <p><strong>Fecha Fin:</strong> ${end}</p>
        `;
        document.getElementById('taskModal').classList.add('active');
    };

    function renderPlaneadorTable(opsData, colorMap) {
        if (!opsData || opsData.length === 0) {
            ganttChartDiv.innerHTML = '<div class="empty-state"><div class="empty-icon">📊</div><h3>Sin datos para mostrar</h3></div>';
            return;
        }

        // 1. Rango de fechas
        let minDateStr = "9999-12-31";
        let maxDateStr = "0000-01-01";
        
        opsData.forEach(op => {
            op.tasks.forEach(task => {
                if (task.start < minDateStr) minDateStr = task.start;
                if (task.end > maxDateStr) maxDateStr = task.end;
            });
        });
        
        if (minDateStr === "9999-12-31") {
            minDateStr = new Date().toISOString().split('T')[0];
            maxDateStr = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString().split('T')[0];
        } else {
            let maxD = new Date(maxDateStr);
            maxD.setDate(maxD.getDate() + 5);
            maxDateStr = maxD.toISOString().split('T')[0];
        }

        const dates = [];
        let currD = new Date(minDateStr);
        const endD = new Date(maxDateStr);
        while (currD <= endD) {
            dates.push(currD.toISOString().split('T')[0]);
            currD.setDate(currD.getDate() + 1);
        }

        // Column configs
        const cols = [
            { id: 'op', title: 'OP' },
            { id: 'desc', title: 'DESCRIPCION' },
            { id: 'cliente', title: 'CLIENTE' },
            { id: 'zona', title: 'ZONA' },
            { id: 'eje', title: 'EJE' },
            { id: 'tipo', title: 'TIPO' },
            { id: 'fecha_fin', title: 'FECHA FIN' }
        ];

        // 2. Construir HTML
        let html = '<div class="table-wrapper" id="tableWrapper"><table class="gantt-table" id="ganttTable">';
        
        // Cabecera
        html += '<thead><tr>';
        cols.forEach((c, i) => {
            if (i <= 2) {
                html += `<th class="frozen-col frozen-idx-${i}">${c.title}</th>`;
            } else {
                html += `<th>${c.title}</th>`;
            }
        });
        
        dates.forEach(d => {
            const parts = d.split('-');
            const displayDate = `${parts[2]}/${parts[1]}`;
            html += `<th class="date-col">${displayDate}</th>`;
        });
        html += '</tr></thead><tbody>';

        // Cuerpo
        opsData.forEach(op => {
            html += '<tr>';
            html += `<td class="frozen-col frozen-idx-0"><b>${op.details.op}</b></td>`;
            html += `<td class="frozen-col frozen-idx-1" title="${op.details.desc}">${op.details.desc}</td>`;
            html += `<td class="frozen-col frozen-idx-2" title="${op.details.cliente}">${op.details.cliente}</td>`;
            html += `<td>${op.details.zona}</td>`;
            html += `<td>${op.details.eje}</td>`;
            html += `<td>${op.details.tipo}</td>`;
            
            // Format fecha_fin DD/MM/YYYY
            let displayFin = '';
            if(op.details.fecha_fin && op.details.fecha_fin !== '9999-12-31'){
                const fparts = op.details.fecha_fin.split('-');
                displayFin = `${fparts[2]}/${fparts[1]}/${fparts[0]}`;
            }
            html += `<td style="text-align:center;"><b>${displayFin}</b></td>`;
            
            // Mapeo de fechas
            const occupiedDates = {};
            op.tasks.forEach(task => {
                const color = (colorMap && colorMap[task.group]) ? colorMap[task.group] : '#3b82f6';
                const textColor = getContrastYIQ(color);
                
                // Escape quotes for onclick handler
                const safeOp = op.details.op.replace(/'/g, "\\'");
                const safeCliente = op.details.cliente.replace(/'/g, "\\'");
                const safeDesc = op.details.desc.replace(/'/g, "\\'");
                
                let safeMachine = task.machine.replace(/'/g, "\\'");
                if (/^0+$/.test(safeMachine)) {
                    safeMachine = '0';
                }
                
                const clickHandler = `showTaskDetails('${safeOp}', '${safeCliente}', '${safeDesc}', '${safeMachine}', '${task.start}', '${task.end}')`;
                
                task.all_dates.forEach(d => {
                    occupiedDates[d] = {
                        machine: safeMachine,
                        color: color,
                        textColor: textColor,
                        clickHandler: clickHandler
                    };
                });
            });

            // Pintar celdas de fecha
            dates.forEach(d => {
                if (occupiedDates[d]) {
                    const info = occupiedDates[d];
                    html += `<td class="filled-cell" style="background-color: ${info.color}; color: ${info.textColor}; cursor: pointer;" onclick="${info.clickHandler}">${info.machine}</td>`;
                } else {
                    html += `<td class="empty-cell"></td>`;
                }
            });

            html += '</tr>';
        });
        html += '</tbody></table></div>';

        ganttChartDiv.innerHTML = html;

        // Apply dynamic sticky positions based on actual rendered widths
        setTimeout(() => {
            const table = document.getElementById('ganttTable');
            if(!table) return;
            const headerRow = table.querySelector('thead tr');
            if(!headerRow) return;
            
            let currentLeft = 0;
            for(let i = 0; i <= 2; i++) {
                const th = headerRow.querySelector(`.frozen-idx-${i}`);
                if(!th) continue;
                
                const width = th.offsetWidth;
                // Get all cells in this column
                const cells = table.querySelectorAll(`.frozen-idx-${i}`);
                cells.forEach(cell => {
                    cell.style.left = currentLeft + 'px';
                });
                currentLeft += width;
            }
        }, 50);
    }
});
