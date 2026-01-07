import React, { useRef, useEffect } from 'react';

export default function PillInput({
    pills,
    onPillsChange,
    placeholder,
    singleValue = false,
    onFocus,
    readOnly = false,
    noWrap = false
}) {
    const inputRef = useRef(null);
    const containerRef = useRef(null);

    const addPill = (text) => {
        if (!text || !text.trim()) return;

        if (singleValue) {
            onPillsChange([text.trim()]);
        } else {
            onPillsChange([...pills, text.trim()]);
        }
    };

    const removePill = (index) => {
        const newPills = pills.filter((_, i) => i !== index);
        onPillsChange(newPills);
    };

    const handleKeyDown = (e) => {
        if (e.key === ',' || e.key === 'Enter' || e.key === ';') {
            e.preventDefault();
            addPill(inputRef.current.value);
            inputRef.current.value = '';
        } else if (e.key === 'Backspace' && inputRef.current.value === '' && pills.length > 0) {
            removePill(pills.length - 1);
        }
    };

    const handleBlur = () => {
        addPill(inputRef.current.value);
        inputRef.current.value = '';
    };

    const handleContainerClick = () => {
        inputRef.current?.focus();
    };

    return (
        <div
            ref={containerRef}
            onClick={readOnly ? undefined : handleContainerClick}
            className={`flex items-center gap-1.5 p-1.5 border border-gray-300 rounded-md ${
                noWrap ? 'overflow-x-auto' : 'flex-wrap'
            } ${readOnly ? 'bg-gray-100 cursor-default' : 'bg-white cursor-text'}`}
        >
            {pills.map((pill, index) => {
                const isParam = pill.startsWith('{') && pill.endsWith('}');
                return (
                    <span
                        key={index}
                        className={`flex items-center rounded-full px-2 py-0.5 text-sm ${
                            isParam
                                ? 'bg-blue-100 text-blue-800'
                                : 'bg-gray-200 text-gray-700'
                        }`}
                    >
                        {pill}
                        {!readOnly && (
                            <span
                                onClick={(e) => {
                                    e.stopPropagation();
                                    removePill(index);
                                }}
                                className="ml-1.5 cursor-pointer font-bold hover:text-red-600"
                            >
                                ×
                            </span>
                        )}
                    </span>
                );
            })}
            {!readOnly && (
                <input
                    ref={inputRef}
                    type="text"
                    className={`flex-grow border-none outline-none p-1 bg-transparent ${
                        noWrap ? 'min-w-[80px]' : 'min-w-[120px]'
                    }`}
                    placeholder={pills.length === 0 ? placeholder : ''}
                    onKeyDown={handleKeyDown}
                    onBlur={handleBlur}
                    onFocus={onFocus}
                />
            )}
        </div>
    );
}
