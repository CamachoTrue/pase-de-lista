// Demostración por rol (/?demo=maestro | director | independiente): la app se abre en otra pestaña tal
// como la ve alguien que entra con ese rol. Todo vive en esa pestaña (sessionStorage) y se borra al
// cerrarla: no toca los datos reales de este dispositivo ni la cuenta, y no se conecta a la nube.
const ROLES = { maestro: 'maestro', director: 'director', independiente: 'maestro' };
const pedido = new URLSearchParams(location.search).get('demo');
export const DEMO = Object.hasOwn(ROLES, pedido || '') ? pedido : null;
export const rolDemo = () => (DEMO ? ROLES[DEMO] : null);
// Dónde se guarda el estado de la app: en la demostración, solo en esta pestaña.
export const almacenApp = () => (DEMO ? sessionStorage : localStorage);
export const sufijoDemo = DEMO ? `.demo.${DEMO}` : '';

// Perfil de ejemplo que ve cada rol en «Mi cuenta».
export const PERFIL_DEMO = {
  director: { nombre: 'Mtra. Laura Gómez', correo: 'direccion@prepaejemplo.mx', rol: 'director', escuela: 'Preparatoria de Ejemplo', clave: 'EJEMPLO', paquete: 'empresarial' },
  maestro: { nombre: 'Prof. Jorge Méndez', correo: 'jorge.mendez@prepaejemplo.mx', rol: 'maestro', escuela: 'Preparatoria de Ejemplo', clave: 'EJEMPLO', paquete: 'empresarial' },
  independiente: { nombre: 'Mtra. Ana Ruiz', correo: 'ana.ruiz@correo.com', rol: 'maestro', escuela: 'Mtra. Ana Ruiz', clave: 'I7K2Q9X', paquete: 'individual' },
};
