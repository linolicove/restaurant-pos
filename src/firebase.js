// src/firebase.js
import { initializeApp } from 'firebase/app';
import { getDatabase, ref, set, onValue } from 'firebase/database';

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
export const rtdb = getDatabase(app);

// Broadcast local state updates up to Firebase Realtime Database
export const syncToCloud = (nodeName, data) => {
  if (data === undefined || data === null) return;
  try {
    const nodeRef = ref(rtdb, `linoli_pos_live/${nodeName}`);
    set(nodeRef, {
      payload: JSON.stringify(data),
      updatedAt: new Date().toISOString()
    });
  } catch (err) {
    console.warn(`Sync error on ${nodeName}:`, err);
  }
};

// Listen in real time for updates made from any other device
export const subscribeToCloud = (nodeName, callback) => {
  const nodeRef = ref(rtdb, `linoli_pos_live/${nodeName}`);
  return onValue(nodeRef, (snapshot) => {
    if (snapshot.exists()) {
      const val = snapshot.val();
      if (val && val.payload) {
        try {
          callback(JSON.parse(val.payload));
        } catch (e) {
          console.error(`Parse error on ${nodeName}:`, e);
        }
      }
    }
  });
};