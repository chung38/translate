import React, { useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Languages, RefreshCw, Loader2 } from 'lucide-react';
import { doc, getDoc } from 'firebase/firestore';
import { db } from '../firebase';
import { User } from 'firebase/auth';

interface UpgradeModalProps {
  isOpen: boolean;
  onClose: () => void;
  user: User | null;
  setStatus: (status: any) => void;
  setStatusMessage: (msg: string) => void;
  setError: (err: string | null) => void;
}

export const UpgradeModal: React.FC<UpgradeModalProps> = ({
  isOpen,
  onClose,
  user,
  setStatus,
  setStatusMessage,
  setError
}) => {
  const [isPaying, setIsPaying] = useState(false);
  const [pendingOrderId, setPendingOrderId] = useState<string | null>(null);

  const checkPaymentStatusManually = async () => {
    if (!pendingOrderId) return;
    setStatus('processing');
    setStatusMessage('正在手動檢查支付狀態...');
    try {
      const paymentRef = doc(db, 'payments', pendingOrderId);
      const paymentDoc = await getDoc(paymentRef);
      if (paymentDoc.exists() && paymentDoc.data().status === 'completed') {
        setPendingOrderId(null);
        setIsPaying(false);
        setStatus('completed');
        setStatusMessage('支付成功！額度已更新。');
      } else {
        setStatusMessage('尚未偵測到支付完成，請稍後再試。');
        setTimeout(() => setStatus('idle'), 3000);
      }
    } catch (err: any) {
      setError(err.message);
      setStatus('error');
    }
  };

  const handleNewebPay = async () => {
    if (!user) return;
    setIsPaying(true);
    try {
      const response = await fetch('/api/newebpay/request', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          userId: user.uid,
          amount: 200,
          productName: '20份檔案翻譯額度',
          email: user.email || '',
        }),
      });
      const data = await response.json();
      if (data.url) {
        setPendingOrderId(data.orderId);
        
        const form = document.createElement('form');
        form.method = 'POST';
        form.action = data.url;
        
        const fields = ['MerchantID', 'TradeInfo', 'TradeSha', 'Version'];
        fields.forEach(field => {
          const input = document.createElement('input');
          input.type = 'hidden';
          input.name = field;
          input.value = data[field];
          form.appendChild(input);
        });
        
        document.body.appendChild(form);
        form.submit();
      } else {
        throw new Error(data.error || '無法發起支付');
      }
    } catch (err: any) {
      setError(err.message);
      setIsPaying(false);
    }
  };

  return (
    <AnimatePresence>
      {isOpen && (
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[110] flex items-center justify-center p-4"
        >
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="bg-white w-full max-w-md rounded-[32px] shadow-2xl overflow-hidden"
          >
            <div className="p-8 text-center">
              <div className="w-20 h-20 bg-emerald-100 rounded-full flex items-center justify-center mx-auto mb-6">
                <Languages className="w-10 h-10 text-emerald-600" />
              </div>
              <h2 className="text-2xl font-bold text-gray-900 mb-2">升級專業版</h2>
              <p className="text-gray-500 mb-8">解鎖更多翻譯額度，享受流暢體驗</p>
              
              <div className="bg-gray-50 rounded-2xl p-6 mb-8 border border-gray-100">
                <div className="flex items-center justify-between mb-4">
                  <span className="text-gray-600 font-medium">方案內容</span>
                  <span className="text-emerald-600 font-bold">20 份檔案額度</span>
                </div>
                <div className="flex items-center justify-between">
                  <span className="text-gray-600 font-medium">價格</span>
                  <span className="text-2xl font-black text-gray-900">$200 <span className="text-sm font-normal text-gray-400">TWD</span></span>
                </div>
              </div>

              <div className="space-y-3">
                {isPaying && (
                  <button 
                    onClick={checkPaymentStatusManually}
                    className="w-full py-4 bg-emerald-50 text-emerald-600 rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-emerald-100 transition-all"
                  >
                    <RefreshCw className="w-5 h-5" />
                    手動檢查支付狀態
                  </button>
                )}
                <button 
                  onClick={handleNewebPay}
                  disabled={isPaying}
                  className="w-full py-4 bg-blue-600 text-white rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-blue-700 transition-all disabled:opacity-50"
                >
                  {isPaying ? (
                    <Loader2 className="w-5 h-5 animate-spin" />
                  ) : (
                    <>
                      藍新金流 快速支付
                    </>
                  )}
                </button>
                <button 
                  onClick={onClose}
                  className="w-full py-4 text-gray-400 font-medium hover:text-gray-600 transition-colors"
                >
                  稍後再說
                </button>
              </div>
            </div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};
