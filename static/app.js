document.addEventListener('DOMContentLoaded', () => {
    const searchInput = document.getElementById('searchInput');
    const clientesBody = document.getElementById('clientesBody');
    const emptyState = document.getElementById('emptyState');
    const clientesTable = document.getElementById('clientesTable');
    
    let clientesData = [];
    let historialData = [];

    // ── Navegación (Sidebar) ────────────────────────────────
    const navItems = document.querySelectorAll('.nav-item');
    const viewSections = document.querySelectorAll('.view-section');

    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            navItems.forEach(nav => nav.classList.remove('active'));
            viewSections.forEach(section => section.classList.add('hidden'));
            
            item.classList.add('active');
            const targetId = item.getAttribute('data-target');
            document.getElementById(targetId).classList.remove('hidden');

            // Cargar datos al entrar a la vista
            if (targetId === 'view-dashboard') loadDashboard();
            if (targetId === 'view-historial') loadHistorial();
        });
    });

    // ── Toggle Sidebar ──────────────────────────────────────
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebar = document.querySelector('.sidebar');
    
    sidebarToggle.addEventListener('click', () => {
        sidebar.classList.toggle('collapsed');
        localStorage.setItem('sidebarCollapsed', sidebar.classList.contains('collapsed'));
    });
    
    if (localStorage.getItem('sidebarCollapsed') === 'true') {
        sidebar.classList.add('collapsed');
    }

    // ── Light Mode Toggle ───────────────────────────────────
    const lightModeToggle = document.getElementById('lightModeToggle');
    
    lightModeToggle.addEventListener('change', (e) => {
        if (e.target.checked) {
            document.body.classList.add('light-mode');
            localStorage.setItem('lightMode', 'true');
        } else {
            document.body.classList.remove('light-mode');
            localStorage.setItem('lightMode', 'false');
        }
    });

    if (localStorage.getItem('lightMode') === 'true') {
        lightModeToggle.checked = true;
        document.body.classList.add('light-mode');
    }

    // ── Outlook Mode Toggle ─────────────────────────────────
    const outlookModeToggle = document.getElementById('outlookModeToggle');
    
    outlookModeToggle.addEventListener('change', (e) => {
        localStorage.setItem('outlookNuevo', e.target.checked ? 'true' : 'false');
    });

    if (localStorage.getItem('outlookNuevo') === 'true') {
        outlookModeToggle.checked = true;
    }

    // ═══════════════════════════════════════════════════════
    // ⭐ FAVORITOS
    // ═══════════════════════════════════════════════════════
    function getFavoritos() {
        try {
            return JSON.parse(localStorage.getItem('favoritos') || '[]');
        } catch { return []; }
    }

    function toggleFavorito(clienteId) {
        let favs = getFavoritos();
        const id = String(clienteId);
        if (favs.includes(id)) {
            favs = favs.filter(f => f !== id);
        } else {
            favs.push(id);
        }
        localStorage.setItem('favoritos', JSON.stringify(favs));
        renderTable(clientesData);
    }

    // ═══════════════════════════════════════════════════════
    // 1. CARGAR CLIENTES
    // ═══════════════════════════════════════════════════════
    // 💀 SKELETON LOADER
    // ═══════════════════════════════════════════════════════
    function showSkeletons() {
        clientesBody.innerHTML = '';
        clientesTable.classList.remove('hidden');
        emptyState.classList.add('hidden');
        for (let i = 0; i < 5; i++) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="width:40px"><div class="skeleton" style="width:20px;height:20px;border-radius:4px;"></div></td>
                <td><div class="skeleton" style="width:${60 + Math.random()*30}%;height:16px;border-radius:4px;"></div></td>
                <td><div class="skeleton" style="width:${40 + Math.random()*40}%;height:16px;border-radius:4px;"></div></td>
                <td class="actions-col"><div class="skeleton" style="width:36px;height:36px;border-radius:6px;"></div></td>
            `;
            clientesBody.appendChild(tr);
        }
    }

    // 1. CARGAR CLIENTES
    // ═══════════════════════════════════════════════════════
    async function loadClientes() {
        showSkeletons();
        try {
            const response = await fetch('/api/clientes');
            clientesData = await response.json();
            renderTable(clientesData);
        } catch (error) {
            console.error("Error al cargar clientes:", error);
        }
    }

    // Boton Refrescar
    const btnRefrescar = document.getElementById('btnRefrescar');
    btnRefrescar.addEventListener('click', async () => {
        const icono = btnRefrescar.querySelector('i');
        icono.classList.add('fa-spin');
        btnRefrescar.disabled = true;
        await loadClientes();
        setTimeout(() => {
            icono.classList.remove('fa-spin');
            btnRefrescar.disabled = false;
        }, 400);
    });

    // ═══════════════════════════════════════════════════════
    // 2. RENDERIZAR TABLA (con favoritos)
    // ═══════════════════════════════════════════════════════
    function renderTable(data) {
        clientesBody.innerHTML = '';
        
        if (data.length === 0) {
            clientesTable.classList.add('hidden');
            emptyState.classList.remove('hidden');
            return;
        }

        clientesTable.classList.remove('hidden');
        emptyState.classList.add('hidden');

        const favs = getFavoritos();

        // Ordenar: favoritos primero
        const sorted = [...data].sort((a, b) => {
            const aFav = favs.includes(String(a.id)) ? 0 : 1;
            const bFav = favs.includes(String(b.id)) ? 0 : 1;
            return aFav - bFav;
        });

        sorted.forEach(cliente => {
            const tr = document.createElement('tr');
            tr.className = 'client-row';
            tr.dataset.id = cliente.id;

            const isFav = favs.includes(String(cliente.id));
            const asuntoTpl = cliente.asunto || '';
            const asuntoPreview = asuntoTpl.replace('[CLIENTE]', cliente.cliente).replace('[PEDIDO]', '...');

            tr.innerHTML = `
                <td style="width: 40px; text-align: center;">
                    <button class="btn-fav ${isFav ? 'active' : ''}" data-id="${cliente.id}" title="Marcar como favorito">
                        <i class="fa-${isFav ? 'solid' : 'regular'} fa-star"></i>
                    </button>
                </td>
                <td class="td-cliente-name">
                    <div class="cliente-name">${cliente.cliente}</div>
                    <div class="cliente-tooltip">
                        <div class="tooltip-row"><strong>Para:</strong> ${cliente.para || '—'}</div>
                        <div class="tooltip-row"><strong>CC:</strong> ${cliente.cc || '—'}</div>
                        <div class="tooltip-row"><strong>Cuerpo:</strong> ${(cliente.cuerpo || '').replace(/<br\s*\/?>/gi, ' ').substring(0, 100)}${(cliente.cuerpo || '').length > 100 ? '...' : ''}</div>
                    </div>
                </td>
                <td>
                    <div class="asunto-preview" title="${asuntoPreview}">${asuntoPreview}</div>
                </td>
                <td class="actions-col">
                    <button class="btn-generar" onclick="generarCorreo(${cliente.id})" title="Generar Correo">
                        <i class="fa-solid fa-paper-plane"></i>
                    </button>
                </td>
            `;

            // Evento favorito
            tr.querySelector('.btn-fav').addEventListener('click', (e) => {
                e.stopPropagation();
                toggleFavorito(cliente.id);
            });

            // Drag & Drop
            setupDragAndDrop(tr, cliente.id);

            clientesBody.appendChild(tr);
        });
    }

    // 3. Filtrar clientes
    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const filtered = clientesData.filter(c => 
            c.cliente.toLowerCase().includes(query) || 
            c.para.toLowerCase().includes(query)
        );
        renderTable(filtered);
    });

    // 4. Drag & Drop
    function setupDragAndDrop(row, clienteId) {
        row.addEventListener('dragover', (e) => {
            e.preventDefault();
            row.classList.add('drag-over');
        });

        row.addEventListener('dragleave', (e) => {
            e.preventDefault();
            row.classList.remove('drag-over');
        });

        row.addEventListener('drop', (e) => {
            e.preventDefault();
            row.classList.remove('drag-over');

            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                // Multi-archivo: pasar TODOS los archivos
                const files = Array.from(e.dataTransfer.files);
                generarCorreo(clienteId, files, row);
            }
        });
    }

    // 5. Generar Correo (soporta multi-archivo)
    window.generarCorreo = async function(clienteId, files = null, row = null) {
        const btn = row ? null : document.querySelector(`tr[data-id="${clienteId}"] .btn-generar`);
        
        if (btn) {
            btn.innerHTML = `<i class="fa-solid fa-spinner"></i>`;
            btn.disabled = true;
        }

        const formData = new FormData();
        formData.append("cliente_id", clienteId);
        const modoOutlook = localStorage.getItem('outlookNuevo') === 'true' ? 'nuevo' : 'clasico';
        formData.append("modo", modoOutlook);
        
        // Soportar un solo archivo o array de archivos
        if (files) {
            const fileList = Array.isArray(files) ? files : [files];
            fileList.forEach(f => formData.append("archivo", f));
        }

        try {
            const response = await fetch('/api/generar_correo', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();
            if (!result.success) {
                alert("Error: " + result.error);
            }
        } catch (error) {
            console.error("Error al generar correo:", error);
            alert("Hubo un error de conexión con el Agente local.");
        } finally {
            if (btn) {
                btn.innerHTML = `<i class="fa-solid fa-paper-plane"></i>`;
                btn.disabled = false;
            }
            if (files && row) {
                row.style.backgroundColor = "var(--success)";
                setTimeout(() => row.style.backgroundColor = "", 1000);
            }
        }
    };

    // ═══════════════════════════════════════════════════════
    // 📊 DASHBOARD
    // ═══════════════════════════════════════════════════════
    async function loadDashboard() {
        try {
            const res = await fetch('/api/stats');
            const stats = await res.json();

            // Animar contadores
            animateCounter('statHoy', stats.hoy);
            animateCounter('statSemana', stats.semana);
            animateCounter('statMes', stats.mes);
            animateCounter('statTotal', stats.total);

            // Top clientes
            const topDiv = document.getElementById('topClientes');
            if (stats.top_clientes.length === 0) {
                topDiv.innerHTML = '<p style="color: var(--text-muted);">Genera correos para ver estadísticas aquí.</p>';
                return;
            }

            const maxCount = stats.top_clientes[0].cantidad;
            topDiv.innerHTML = stats.top_clientes.map((c, i) => `
                <div class="top-client-item">
                    <span class="top-client-rank">#${i + 1}</span>
                    <div class="top-client-info">
                        <div class="top-client-name">${c.nombre}</div>
                        <div class="top-client-bar-bg">
                            <div class="top-client-bar" style="width: ${(c.cantidad / maxCount * 100)}%"></div>
                        </div>
                    </div>
                    <span class="top-client-count">${c.cantidad}</span>
                </div>
            `).join('');
        } catch (error) {
            console.error("Error al cargar dashboard:", error);
        }
    }

    function animateCounter(elementId, target) {
        const el = document.getElementById(elementId);
        const current = parseInt(el.textContent) || 0;
        if (current === target) return;

        const duration = 600;
        const start = performance.now();

        function step(timestamp) {
            const progress = Math.min((timestamp - start) / duration, 1);
            const eased = 1 - Math.pow(1 - progress, 3); // ease-out cubic
            el.textContent = Math.round(current + (target - current) * eased);
            if (progress < 1) requestAnimationFrame(step);
        }
        requestAnimationFrame(step);
    }

    // ═══════════════════════════════════════════════════════
    // 📝 HISTORIAL
    // ═══════════════════════════════════════════════════════
    const historialSearch = document.getElementById('historialSearch');
    const historialFiltroFecha = document.getElementById('historialFiltroFecha');

    async function loadHistorial() {
        try {
            const res = await fetch('/api/historial');
            historialData = await res.json();
            renderHistorial();
        } catch (error) {
            console.error("Error al cargar historial:", error);
        }
    }

    function renderHistorial() {
        const body = document.getElementById('historialBody');
        const table = document.getElementById('historialTable');
        const empty = document.getElementById('historialEmpty');

        // Filtrar por texto
        const query = (historialSearch?.value || '').toLowerCase();
        const filtroFecha = historialFiltroFecha?.value || 'todos';
        const ahora = new Date();
        const hoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());

        let filtered = historialData.filter(entry => {
            // Filtro texto
            const matchTexto = !query || 
                (entry.cliente || '').toLowerCase().includes(query) ||
                (entry.asunto || '').toLowerCase().includes(query);
            
            // Filtro fecha
            let matchFecha = true;
            if (filtroFecha !== 'todos') {
                const entryDate = new Date(entry.timestamp);
                if (filtroFecha === 'hoy') {
                    matchFecha = entryDate >= hoy;
                } else if (filtroFecha === 'semana') {
                    const inicioSemana = new Date(hoy);
                    inicioSemana.setDate(hoy.getDate() - hoy.getDay() + 1); // Lunes
                    matchFecha = entryDate >= inicioSemana;
                } else if (filtroFecha === 'mes') {
                    const inicioMes = new Date(ahora.getFullYear(), ahora.getMonth(), 1);
                    matchFecha = entryDate >= inicioMes;
                }
            }

            return matchTexto && matchFecha;
        });

        body.innerHTML = '';

        if (filtered.length === 0) {
            table.classList.add('hidden');
            empty.classList.remove('hidden');
            return;
        }

        table.classList.remove('hidden');
        empty.classList.add('hidden');

        filtered.forEach(entry => {
            const tr = document.createElement('tr');
            const fecha = new Date(entry.timestamp);
            const fechaStr = fecha.toLocaleDateString('es-CL', { day: '2-digit', month: '2-digit', year: 'numeric' }) 
                           + ' ' + fecha.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit' });
            
            const modoClass = entry.modo === 'clasico' ? 'badge-clasico' : 'badge-nuevo';
            const modoLabel = entry.modo === 'clasico' ? 'Clásico' : 'Nuevo';
            const adjuntoHtml = entry.archivo 
                ? `<span class="badge-adjunto"><i class="fa-solid fa-paperclip"></i> ${entry.archivo}</span>`
                : `<span style="color: var(--text-muted); font-size: 0.85rem;">—</span>`;

            tr.innerHTML = `
                <td style="font-size: 0.85rem; white-space: nowrap;">${fechaStr}</td>
                <td><span class="cliente-name" style="font-size: 0.9rem;">${entry.cliente || ''}</span></td>
                <td><span style="color: var(--text-muted); font-size: 0.85rem;">${entry.asunto || ''}</span></td>
                <td>${adjuntoHtml}</td>
                <td><span class="badge-modo ${modoClass}">${modoLabel}</span></td>
            `;
            body.appendChild(tr);
        });
    }

    // Eventos de filtro del historial
    if (historialSearch) {
        historialSearch.addEventListener('input', renderHistorial);
    }
    if (historialFiltroFecha) {
        historialFiltroFecha.addEventListener('change', renderHistorial);
    }

    // ═══════════════════════════════════════════════════════
    // 🔄 AUTO-RECARGA EXCEL
    // ═══════════════════════════════════════════════════════
    let lastExcelTimestamp = 0;

    async function checkExcelChanges() {
        try {
            const res = await fetch('/api/excel_timestamp');
            const data = await res.json();
            if (lastExcelTimestamp > 0 && data.timestamp > lastExcelTimestamp) {
                // Excel cambió — recargar
                await loadClientes();
                // Flash sutil en el botón de refrescar
                const icono = btnRefrescar.querySelector('i');
                icono.classList.add('fa-spin');
                setTimeout(() => icono.classList.remove('fa-spin'), 600);
            }
            lastExcelTimestamp = data.timestamp;
        } catch (error) {
            // Silencioso — no interrumpir
        }
    }

    // Polling cada 10 segundos
    setInterval(checkExcelChanges, 10000);
    // Obtener timestamp inicial
    checkExcelChanges();

    // ═══════════════════════════════════════════════════════
    // BOTÓN ABRIR EXCEL
    // ═══════════════════════════════════════════════════════
    const btnAbrirExcel = document.getElementById('btnAbrirExcel');
    if (btnAbrirExcel) {
        btnAbrirExcel.addEventListener('click', async () => {
            try {
                btnAbrirExcel.innerHTML = `<i class="fa-solid fa-spinner"></i> Abriendo...`;
                await fetch('/api/abrir_excel');
            } catch (error) {
                console.error(error);
            } finally {
                setTimeout(() => {
                    btnAbrirExcel.innerHTML = `<i class="fa-solid fa-file-excel"></i> Abrir Excel`;
                }, 1000);
            }
        });
    }

    // ═══════════════════════════════════════════════════════
    // 📤 EXPORTAR HISTORIAL
    // ═══════════════════════════════════════════════════════
    const btnExportar = document.getElementById('btnExportarHistorial');
    if (btnExportar) {
        btnExportar.addEventListener('click', () => {
            btnExportar.innerHTML = `<i class="fa-solid fa-spinner"></i> Exportando...`;
            btnExportar.disabled = true;
            // Descargar archivo
            const link = document.createElement('a');
            link.href = '/api/exportar_historial';
            link.download = 'Historial_Correos.xlsx';
            link.click();
            setTimeout(() => {
                btnExportar.innerHTML = `<i class="fa-solid fa-file-export"></i> Exportar Excel`;
                btnExportar.disabled = false;
            }, 1500);
        });
    }

    // Iniciar
    loadClientes();
});
