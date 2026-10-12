import { initializeApp } from "firebase/app";
import { getDatabase, ref, set, push, onValue, off } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyClbjQD6DEcmYtLx6IbOehZdmjt016Gos8",
  authDomain: "lc-pos-systems.firebaseapp.com",
  databaseURL: "https://lc-pos-systems-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "lc-pos-systems",
  storageBucket: "lc-pos-systems.firebasestorage.app",
  messagingSenderId: "821657587922",
  appId: "1:821657587922:web:c23231fce32da446435d0b",
  measurementId: "G-6ESVVZNNQE"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Realtime Database
export const rtdb = getDatabase(app);

/**
 * Syncs local state directly to pos_state/{node}
 * @param {string} node - Target node path (e.g. 'active_orders', 'menu_items')
 * @param {any} data - State value to store
 */
export const syncToCloud = async (node, data) => {
  if (data === undefined) return;
  try {
    const nodeRef = ref(rtdb, `pos_state/${node}`);
    await set(nodeRef, data);
  } catch (error) {
    console.warn(`RTDB sync error on [${node}]:`, error);
  }
};

/**
 * Appends an immutable, permanent historical record to pos_archives/{path}
 * Uses atomic push() so data is never overwritten or deleted.
 * @param {string} path - Target archive node (e.g. 'attendance_logs', 'payroll_records')
 * @param {object} record - Data payload to archive forever
 */
export const appendCloudArchive = async (path, record) => {
  if (!record) return;
  try {
    const listRef = ref(rtdb, `pos_archives/${path}`);
    const newRecordRef = push(listRef);
    await set(newRecordRef, {
      ...record,
      serverArchivedAt: new Date().toISOString()
    });
  } catch (error) {
    console.warn(`RTDB archive append error on [${path}]:`, error);
  }
};

/**
 * Listens for real-time updates from other terminals and devices
 * @param {string} node - Node path to listen to
 * @param {function} callback - Function to receive incoming data
 * @returns {function} Unsubscribe callback
 */
export const subscribeToCloud = (node, callback) => {
  try {
    const nodeRef = ref(rtdb, `pos_state/${node}`);
    const listener = onValue(
      nodeRef,
      (snapshot) => {
        if (snapshot.exists()) {
          callback(snapshot.val());
        }
      },
      (error) => {
        console.warn(`RTDB listener error on [${node}]:`, error);
      }
    );

    return () => off(nodeRef, "value", listener);
  } catch (error) {
    console.warn(`Failed to attach RTDB listener on [${node}]:`, error);
    return () => {};
  }
};