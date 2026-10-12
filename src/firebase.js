import { initializeApp } from "firebase/app";
import { getDatabase, ref, set, onValue } from "firebase/database";

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

const app = initializeApp(firebaseConfig);

// Regional instance URL must be passed explicitly
export const rtdb = getDatabase(app, "https://lc-pos-systems-default-rtdb.asia-southeast1.firebasedatabase.app");

export const syncToCloud = async (nodeKey, payload) => {
  try {
    const nodeRef = ref(rtdb, `pos_state/${nodeKey}`);
    await set(nodeRef, payload);
    console.log(`[Firebase] Successfully synced: pos_state/${nodeKey}`);
  } catch (err) {
    console.error(`[Firebase] Sync failed for pos_state/${nodeKey}:`, err);
  }
};

export const subscribeToCloud = (nodeKey, callback) => {
  try {
    const nodeRef = ref(rtdb, `pos_state/${nodeKey}`);
    return onValue(nodeRef, (snapshot) => {
      const val = snapshot.val();
      callback(val);
    }, (err) => {
      console.error(`[Firebase] Subscription error on pos_state/${nodeKey}:`, err);
    });
  } catch (err) {
    console.error(`[Firebase] Error initializing listener:`, err);
    return () => {};
  }
};

export default app;