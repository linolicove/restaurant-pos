import { initializeApp } from "firebase/app";
import { getDatabase, ref, set, onValue, off } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyAbSLBjSFjlFdAG-8_oqoIMI1DkyP1Aew4",
  authDomain: "restaurant-pos-3239a.firebaseapp.com",
  projectId: "restaurant-pos-3239a",
  storageBucket: "restaurant-pos-3239a.firebasestorage.app",
  messagingSenderId: "232963602696",
  appId: "1:232963602696:web:6f9779306e192cbf3654bd",
  measurementId: "G-8KK29854B3"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Realtime Database
export const rtdb = getDatabase(app);

/**
 * Syncs local state changes directly to Firebase Realtime Database
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

    return () => off(nodeRef, 'value', listener);
  } catch (error) {
    console.warn(`Failed to attach RTDB listener on [${node}]:`, error);
    return () => {};
  }
};