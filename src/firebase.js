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

// Initialize App
const app = initializeApp(firebaseConfig);

// Initialize RTDB with exact regional URL
export const rtdb = getDatabase(app, "https://lc-pos-systems-default-rtdb.asia-southeast1.firebasedatabase.app");

// Push any collection or object up to Firebase
export const syncToCloud = (nodeKey, payload) => {
  try {
    const nodeRef = ref(rtdb, `pos_state/${nodeKey}`);
    return set(nodeRef, payload);
  } catch (err) {
    console.error(`Firebase sync error for ${nodeKey}:`, err);
  }
};

// Listen in real-time to changes from Firebase
export const subscribeToCloud = (nodeKey, callback) => {
  try {
    const nodeRef = ref(rtdb, `pos_state/${nodeKey}`);
    return onValue(nodeRef, (snapshot) => {
      const val = snapshot.val();
      if (val !== null && val !== undefined) {
        callback(val);
      }
    });
  } catch (err) {
    console.error(`Firebase subscribe error for ${nodeKey}:`, err);
    return () => {};
  }
};

export default app;