import React from "react";

/**
 * Media - simple responsive image/GIF component.
 * Props:
 *  - src: string (required) image/GIF URL or import
 *  - alt: string
 *  - width: string|number (default "100%")
 *  - height: string|number (default "auto")
 *  - fit: "contain" | "cover" | "fill" (object-fit)
 *  - clickable: boolean (if true opens image in new tab on click)
 *  - style: additional wrapper styles
 */
export default function Media({ src, alt = "", width = "100%", height = "auto", fit = "contain", clickable = false, style = {}, className = "" }) {
	const img = (
		<img
			src={src}
			alt={alt}
			loading="lazy"
			style={{
				display: "block",
				maxWidth: "100%",
				width,
				height,
				objectFit: fit,
			}}
		/>
	);

	return (
		<div
			className={className}
			style={{
				display: "flex",
				justifyContent: "center",
				alignItems: "center",
				overflow: "hidden",
				...style,
			}}
			onClick={() => {
				if (clickable && src) window.open(src, "_blank", "noopener");
			}}
			role={clickable ? "button" : undefined}
		>
			{img}
		</div>
	);
}

/**
 * Gif helper — shorthand to insert a GIF.
 * Usage: <Gif src="https://example.com/anim.gif" alt="demo" width="400px" clickable />
 */
export const Gif = (props) => {
	return <Media {...props} />;
};
