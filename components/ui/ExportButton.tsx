"use client"

import { Button } from "@/components/ui/button"
import { FileDown, Loader2 } from "lucide-react"
import { useState } from "react"
import {
    Tooltip,
    TooltipContent,
    TooltipProvider,
    TooltipTrigger,
} from "@/components/ui/tooltip"

interface ExportButtonProps {
    onExport: () => Promise<void>;
    label?: string;
    showLabel?: boolean;
    variant?: "default" | "outline" | "ghost";
    size?: "default" | "sm" | "lg" | "icon";
    className?: string;
}

/**
 * Reusable Export to Excel button component
 * Matches the wine analytics dashboard theme (black, neutral, purple, red)
 */
export default function ExportButton({
    onExport,
    label = "Export to Excel",
    showLabel = false,
    variant = "outline",
    size = "sm",
    className = ""
}: ExportButtonProps) {
    const [isExporting, setIsExporting] = useState(false);

    const handleExport = async () => {
        try {
            setIsExporting(true);
            await onExport();
        } catch (error) {
            console.error("Export failed:", error);
            // You could add toast notification here if you have a toast system
        } finally {
            setIsExporting(false);
        }
    };

    const buttonContent = (
        <Button
            onClick={handleExport}
            disabled={isExporting}
            variant={variant}
            size={size}
            className={`
                transition-all duration-200
                hover:bg-neutral-800 hover:text-white
                border-neutral-700
                ${isExporting ? 'opacity-50 cursor-not-allowed' : ''}
                ${className}
            `}
        >
            {isExporting ? (
                <>
                    <Loader2 className="h-4 w-4 animate-spin" />
                    {showLabel && <span className="ml-2">Exporting...</span>}
                </>
            ) : (
                <>
                    <FileDown className="h-4 w-4" />
                    {showLabel && <span className="ml-2">{label}</span>}
                </>
            )}
        </Button>
    );

    // If label is hidden, wrap in tooltip
    if (!showLabel) {
        return (
            <TooltipProvider>
                <Tooltip>
                    <TooltipTrigger asChild>
                        {buttonContent}
                    </TooltipTrigger>
                    <TooltipContent>
                        <p>{label}</p>
                    </TooltipContent>
                </Tooltip>
            </TooltipProvider>
        );
    }

    return buttonContent;
}