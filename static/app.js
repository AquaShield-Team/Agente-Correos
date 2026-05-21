document.addEventListener('DOMContentLoaded', () => {
    const searchInput = document.getElementById('searchInput');
    const clientesBody = document.getElementById('clientesBody');
    const emptyState = document.getElementById('emptyState');
    const clientesTable = document.getElementById('clientesTable');
    
    let clientesData = [];

    // Navegación (Sidebar)
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
        });
    });

    // Toggle Sidebar
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebar = document.querySelector('.sidebar');
    
    sidebarToggle.addEventListener('click', () => {
        sidebar.classList.toggle('collapsed');
        // Guardar preferencia
        localStorage.setItem('sidebarCollapsed', sidebar.classList.contains('collapsed'));
    });
    
    // Restaurar preferencia del sidebar
    if (localStorage.getItem('sidebarCollapsed') === 'true') {
        sidebar.classList.add('collapsed');
    }

    // Light Mode Toggle
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

    // Restaurar preferencia de Light Mode
    if (localStorage.getItem('lightMode') === 'true') {
        lightModeToggle.checked = true;
        document.body.classList.add('light-mode');
    }

    // Outlook Mode Toggle (clasico vs nuevo)
    const outlookModeToggle = document.getElementById('outlookModeToggle');
    
    outlookModeToggle.addEventListener('change', (e) => {
        localStorage.setItem('outlookNuevo', e.target.checked ? 'true' : 'false');
    });

    // Restaurar preferencia de Outlook
    if (localStorage.getItem('outlookNuevo') === 'true') {
        outlookModeToggle.checked = true;
    }

    // 1. Cargar datos del servidor
    async function loadClientes() {
        try {
            const response = await fetch('/api/clientes');
            clientesData = await response.json();
            renderTable(clientesData);
        } catch (error) {
            console.error("Error al cargar clientes:", error);
            alert("Error al cargar la base de datos de clientes.");
        }
    }

    // Boton Refrescar: recarga la lista desde el Excel
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

    // 2. Renderizar tabla
    function renderTable(data) {
        clientesBody.innerHTML = '';
        
        if (data.length === 0) {
            clientesTable.classList.add('hidden');
            emptyState.classList.remove('hidden');
            return;
        }

        clientesTable.classList.remove('hidden');
        emptyState.classList.add('hidden');

        data.forEach(cliente => {
            const tr = document.createElement('tr');
            tr.className = 'client-row';
            tr.dataset.id = cliente.id;

            // Preparar el asunto preview
            const asuntoTpl = cliente.asunto || '';
            const asuntoPreview = asuntoTpl.replace('[CLIENTE]', cliente.cliente).replace('[PEDIDO]', '...');

            tr.innerHTML = `
                <td>
                    <div class="cliente-name">${cliente.cliente}</div>
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

            // Agregar eventos de Drag & Drop
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

    // 4. Lógica de Drag & Drop por fila
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
                const file = e.dataTransfer.files[0];
                generarCorreo(clienteId, file, row);
            }
        });
    }

    // 5. Función principal para generar correo
    window.generarCorreo = async function(clienteId, file = null, row = null) {
        const btn = row ? null : document.querySelector(`tr[data-id="${clienteId}"] .btn-generar`);
        
        if (btn) {
            btn.innerHTML = `<i class="fa-solid fa-spinner"></i>`;
            btn.disabled = true;
        }

        const formData = new FormData();
        formData.append("cliente_id", clienteId);
        // Enviar modo de Outlook al backend
        const modoOutlook = localStorage.getItem('outlookNuevo') === 'true' ? 'nuevo' : 'clasico';
        formData.append("modo", modoOutlook);
        if (file) {
            formData.append("archivo", file);
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
            if (file) {
                // Pequeña notificación visual de éxito en la fila
                row.style.backgroundColor = "var(--success)";
                setTimeout(() => row.style.backgroundColor = "", 1000);
            }
        }
    };


    // Botón Abrir Excel
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
                    btnAbrirExcel.innerHTML = `<i class="fa-solid fa-file-excel"></i> Abrir Excel de Base de Datos`;
                }, 1000);
            }
        });
    }

    // 7. Añadir cliente (Formulario)
    const formNuevoCliente = document.getElementById('formNuevoCliente');
    if (formNuevoCliente) {
        formNuevoCliente.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            const payload = {
                cliente: document.getElementById('new_cliente').value,
                para: document.getElementById('new_para').value,
                cc: document.getElementById('new_cc').value,
                asunto: document.getElementById('new_asunto').value,
                cuerpo: document.getElementById('new_cuerpo').value,
            };

            const btn = formNuevoCliente.querySelector('button');
            btn.innerHTML = `<i class="fa-solid fa-spinner"></i> Guardando...`;
            btn.disabled = true;

            try {
                const res = await fetch('/api/clientes', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const result = await res.json();
                if (result.success) {
                    formNuevoCliente.reset();
                    loadClientes();
                    // Volver a la vista de correos
                    document.querySelector('[data-target="view-correos"]').click();
                }
            } catch (error) {
                console.error(error);
                alert("Error al guardar el cliente.");
            } finally {
                btn.innerHTML = `<i class="fa-solid fa-plus"></i> Guardar Cliente`;
                btn.disabled = false;
            }
        });
    }

    // Iniciar
    loadClientes();
});
