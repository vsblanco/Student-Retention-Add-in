import { render, screen } from '@testing-library/react';
import StudentHeader from './StudentHeader';

test('renders StudentHeader component', () => {
    render(<StudentHeader />);
    const headerElement = screen.getByText(/Student Header/i);
    expect(headerElement).toBeInTheDocument();
});