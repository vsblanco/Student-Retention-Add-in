import React from 'react';

function BounceAnimation() {
  return (
    <style>
      {`
        .bounce {
          animation: bounce 0.5s;
        }
        @keyframes bounce {
          0%   { transform: scale(1); }
          20%  { transform: scale(1.15); }
          40%  { transform: scale(0.95); }
          60%  { transform: scale(1.10); }
          80%  { transform: scale(0.98); }
          100% { transform: scale(1); }
        }
      `}
    </style>
  );
}

export default BounceAnimation;
